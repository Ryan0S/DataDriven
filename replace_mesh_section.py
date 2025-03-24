import re
import shutil
import os
import xml.etree.ElementTree as ET
import zipfile
import tempfile
import json
from typing import List, Dict, Optional

# === CONFIG ===
CONFIG = {
    "template_3mf_file": r"C:\Users\Ryan\Downloads\NumberedBonesPre.3mf",
    "output_3mf_folder_base": r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified",
    "output_3mf_file_base": r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified",
    "max_objects_per_file": 6,  # Maximum number of objects per output file
    "objects_json_file": r"C:\Users\Ryan\Desktop\Tidy\DataDriven\objects.json"  # Path to the JSON file containing objects
}

# Dynamic source file path generator
SOURCE_FILE_BASE_DIR = r"C:\Users\Ryan\Downloads\DogboneExports"
def get_source_file_path(specimen_id: str) -> str:
    """Generate source file path dynamically based on specimen_id."""
    padded_id = specimen_id.zfill(3)
    return os.path.join(SOURCE_FILE_BASE_DIR, f"{padded_id}.3mf")

# === UTILITY FUNCTIONS ===
def extract_3mf(archive_path: str, extract_to: str) -> str:
    """Extract a .3mf file to a specified directory."""
    if not os.path.exists(archive_path):
        raise FileNotFoundError(f"❌ Source file not found: {archive_path}")
    with zipfile.ZipFile(archive_path, 'r') as z:
        z.extractall(extract_to)
    print(f"Extracted {archive_path} to {extract_to}")
    return extract_to

def rezip_3mf(folder_path: str, output_path: str) -> None:
    """Repack a folder into a .3mf file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder_path):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, folder_path)
                z.write(abs_path, rel_path)
    print(f"✅ Repacked into: {output_path}")

def read_file(file_path: str) -> str:
    """Read content from a file with UTF-8 encoding."""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def write_file(file_path: str, content: str) -> None:
    """Write content to a file with UTF-8 encoding."""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

def load_objects_from_json(json_file: str) -> List[Dict[str, str]]:
    """Load the objects list from a JSON file."""
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            objects = json.load(f)
        # Validate the structure
        for obj in objects:
            if not all(key in obj for key in ["specimen_id", "fill_density", "fill_pattern"]):
                raise ValueError(f"❌ Invalid object in JSON: {obj}. Must include specimen_id, fill_density, and fill_pattern.")
        return objects
    except FileNotFoundError:
        raise FileNotFoundError(f"❌ JSON file not found: {json_file}")
    except json.JSONDecodeError:
        raise ValueError(f"❌ Invalid JSON format in file: {json_file}")

# === 3MF PROCESSING ===
class ModelProcessor:
    def __init__(self, template_folder: str, output_folder: str):
        self.template_path = os.path.join(template_folder, "3D", "3dmodel.model")
        self.template_folder = template_folder
        self.output_folder = output_folder
        self.output_model_path = os.path.join(output_folder, "3D", "3dmodel.model")
        self.output_config_path = os.path.join(output_folder, "Metadata", "Slic3r_PE_model.config")

    def setup_output_folder(self) -> None:
        """Duplicate the entire template folder to output, replacing existing if necessary."""
        if os.path.exists(self.output_folder):
            shutil.rmtree(self.output_folder)
        shutil.copytree(self.template_folder, self.output_folder)

    def extract_mesh(self, text: str) -> str:
        """Extract mesh section from model text."""
        match = re.search(r"<mesh>.*?</mesh>", text, re.DOTALL)
        if not match:
            raise ValueError("❌ Could not find <mesh> section in source file.")
        return match.group(0)

    def count_triangles(self, text: str) -> int:
        """Count triangles in model text."""
        return len(re.findall(r"<triangle\b", text)) - 1

    def replace_mesh(self, source_text: str, target_text: str, object_id: str) -> str:
        """Replace mesh in target text for a specific object ID."""
        pattern = fr'(<object id="{object_id}" type="model">.*?)(<mesh>.*?</mesh>)(.*?</object>)'
        match = re.search(pattern, target_text, re.DOTALL)
        if not match:
            raise ValueError(f"❌ Could not find object id={object_id} with a <mesh> in target file.")
        
        before, _, after = match.groups()
        new_mesh = self.extract_mesh(source_text)
        replaced_object = before + new_mesh + after
        return target_text.replace(match.group(0), replaced_object)

    def update_config(self, objects: List[Dict[str, str]], source_paths: Dict[str, str]) -> None:
        """Update config file for multiple objects with triangle count, fill density, and pattern."""
        tree = ET.parse(self.output_config_path)
        root = tree.getroot()

        for index, obj_config in enumerate(objects, start=1):  # Start IDs from 1
            object_id = str(index)
            fill_density = obj_config["fill_density"]
            fill_pattern = obj_config["fill_pattern"]
            specimen_id = obj_config["specimen_id"]
            source_path = os.path.join(source_paths[specimen_id], "3D", "3dmodel.model")
            source_text = read_file(source_path)
            triangle_count = self.count_triangles(source_text)

            obj = root.find(f".//object[@id='{object_id}']")
            if obj is None:
                raise ValueError(f"❌ Could not find object id={object_id} in config file.")

            self._update_metadata(obj, "fill_density", fill_density)
            self._update_metadata(obj, "fill_pattern", fill_pattern)
            
            volume = obj.find("volume")
            if volume is None:
                raise ValueError(f"❌ Could not find volume for object id={object_id} in config file.")
            volume.set("lastid", str(triangle_count))

        tree.write(self.output_config_path, encoding="utf-8", xml_declaration=True)

    def _update_metadata(self, obj: ET.Element, key: str, value: str) -> None:
        """Update or create a metadata element in the object."""
        elem = obj.find(f".//metadata[@key='{key}']")
        if elem is not None:
            elem.set("value", value)
        else:
            ET.SubElement(obj, "metadata", {"type": "object", "key": key, "value": value})

    def process(self, objects: List[Dict[str, str]], source_paths: Dict[str, str]) -> str:
        """Process the 3MF model for a group of objects and return the output folder."""
        self.setup_output_folder()
        
        target_text = read_file(self.template_path)
        
        # Replace mesh for each object using its specific source file and dynamic object_id
        new_text = target_text
        for index, obj_config in enumerate(objects, start=1):  # Start IDs from 1
            object_id = str(index)
            specimen_id = obj_config["specimen_id"]
            source_path = os.path.join(source_paths[specimen_id], "3D", "3dmodel.model")
            source_text = read_file(source_path)
            new_text = self.replace_mesh(source_text, new_text, object_id)
        
        write_file(self.output_model_path, new_text)
        
        # Update config for all objects in this group
        self.update_config(objects, source_paths)
        
        for index, obj_config in enumerate(objects, start=1):
            object_id = str(index)
            specimen_id = obj_config["specimen_id"]
            source_path = os.path.join(source_paths[specimen_id], "3D", "3dmodel.model")
            triangle_count = self.count_triangles(read_file(source_path))
            print(f"✅ Mesh replaced, config updated (fill_density={obj_config['fill_density']}, "
                  f"fill_pattern={obj_config['fill_pattern']}, lastid={triangle_count}) "
                  f"for object id={object_id} from specimen {specimen_id} in:\n{self.output_model_path}")
        return self.output_folder

# === MAIN ===
def main() -> None:
    # Load objects from JSON file
    objects = load_objects_from_json(CONFIG["objects_json_file"])
    
    # Extract all source files based on specimen_ids
    source_paths = {}
    with tempfile.TemporaryDirectory() as temp_template_dir:
        template_folder = extract_3mf(CONFIG["template_3mf_file"], temp_template_dir)
        
        # Create temporary directories for each unique specimen_id
        temp_dirs = {}
        try:
            for obj_config in objects:
                specimen_id = obj_config["specimen_id"]
                if specimen_id not in temp_dirs:
                    source_file = get_source_file_path(specimen_id)
                    temp_dirs[specimen_id] = tempfile.TemporaryDirectory()
                    source_paths[specimen_id] = extract_3mf(source_file, temp_dirs[specimen_id].name)

            # Process objects in groups of max_objects_per_file
            max_objects = CONFIG["max_objects_per_file"]
            total_objects = len(objects)
            num_files = (total_objects + max_objects - 1) // max_objects  # Ceiling division

            for file_idx in range(num_files):
                start_idx = file_idx * max_objects
                end_idx = min(start_idx + max_objects, total_objects)
                group_objects = objects[start_idx:end_idx]

                # Generate unique output folder and file names
                output_folder = f"{CONFIG['output_3mf_folder_base']}_{file_idx + 1}"
                output_file = f"{CONFIG['output_3mf_file_base']}_{file_idx + 1}.3mf"

                processor = ModelProcessor(template_folder, output_folder)
                processor.process(group_objects, source_paths)
                rezip_3mf(output_folder, output_file)

        finally:
            # Clean up temporary directories
            for temp_dir in temp_dirs.values():
                temp_dir.cleanup()

if __name__ == "__main__":
    main()