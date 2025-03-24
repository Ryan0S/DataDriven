import re
import shutil
import os
import xml.etree.ElementTree as ET
import zipfile
import tempfile
from typing import Optional

# === CONFIG ===
CONFIG = {
    "source_3mf_file": r"C:\Users\Ryan\Downloads\Dogbone_7.3mf",
    "template_3mf_file": r"C:\Users\Ryan\Downloads\NumberedBonesPre.3mf",
    "output_3mf_folder": r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified",
    "output_3mf_file": r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified.3mf",
    "object_id": "1",
    "fill_density": "70%",
    "fill_pattern": "grid",
}

# === UTILITY FUNCTIONS ===
def extract_3mf(archive_path: str, extract_to: str) -> str:
    """Extract a .3mf file to a specified directory."""
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

# === 3MF PROCESSING ===
class ModelProcessor:
    def __init__(self, source_folder: str, template_folder: str, output_folder: str, object_id: str):
        self.source_path = os.path.join(source_folder, "3D", "3dmodel.model")
        self.template_path = os.path.join(template_folder, "3D", "3dmodel.model")  # Added template_path
        self.template_folder = template_folder
        self.output_folder = output_folder
        self.output_model_path = os.path.join(output_folder, "3D", "3dmodel.model")
        self.output_config_path = os.path.join(output_folder, "Metadata", "Slic3r_PE_model.config")
        self.object_id = object_id

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

    def replace_mesh(self, source_text: str, target_text: str) -> str:
        """Replace mesh in target text with source mesh."""
        pattern = fr'(<object id="{self.object_id}" type="model">.*?)(<mesh>.*?</mesh>)(.*?</object>)'
        match = re.search(pattern, target_text, re.DOTALL)
        if not match:
            raise ValueError(f"❌ Could not find object id={self.object_id} with a <mesh> in target file.")
        
        before, _, after = match.groups()
        new_mesh = self.extract_mesh(source_text)
        replaced_object = before + new_mesh + after
        return target_text.replace(match.group(0), replaced_object)

    def update_config(self, triangle_count: int, fill_density: str, fill_pattern: str) -> None:
        """Update config file with new triangle count, fill density, and pattern."""
        tree = ET.parse(self.output_config_path)
        root = tree.getroot()
        
        obj = root.find(f".//object[@id='{self.object_id}']")
        if obj is None:
            raise ValueError(f"❌ Could not find object id={self.object_id} in config file.")

        self._update_metadata(obj, "fill_density", fill_density)
        self._update_metadata(obj, "fill_pattern", fill_pattern)
        
        volume = obj.find("volume")
        if volume is None:
            raise ValueError(f"❌ Could not find volume for object id={self.object_id} in config file.")
        volume.set("lastid", str(triangle_count))

        tree.write(self.output_config_path, encoding="utf-8", xml_declaration=True)

    def _update_metadata(self, obj: ET.Element, key: str, value: str) -> None:
        """Update or create a metadata element in the object."""
        elem = obj.find(f".//metadata[@key='{key}']")
        if elem is not None:
            elem.set("value", value)
        else:
            ET.SubElement(obj, "metadata", {"type": "object", "key": key, "value": value})

    def process(self, fill_density: str, fill_pattern: str) -> str:
        """Process the 3MF model and return the output folder."""
        self.setup_output_folder()
        
        source_text = read_file(self.source_path)
        target_text = read_file(self.template_path)  # Fixed: Read from template_path, not source_path
        
        new_text = self.replace_mesh(source_text, target_text)
        write_file(self.output_model_path, new_text)
        
        triangle_count = self.count_triangles(source_text)
        self.update_config(triangle_count, fill_density, fill_pattern)
        
        print(f"✅ Mesh replaced, config updated (fill_density={fill_density}, fill_pattern={fill_pattern}, lastid={triangle_count}) "
              f"for object id={self.object_id} in:\n{self.output_model_path}")
        return self.output_folder

# === MAIN ===
def main() -> None:
    with tempfile.TemporaryDirectory() as temp_source_dir, tempfile.TemporaryDirectory() as temp_template_dir:
        source_folder = extract_3mf(CONFIG["source_3mf_file"], temp_source_dir)
        template_folder = extract_3mf(CONFIG["template_3mf_file"], temp_template_dir)

        processor = ModelProcessor(source_folder, template_folder, CONFIG["output_3mf_folder"], CONFIG["object_id"])
        modified_folder = processor.process(CONFIG["fill_density"], CONFIG["fill_pattern"])
        
        rezip_3mf(modified_folder, CONFIG["output_3mf_file"])

if __name__ == "__main__":
    main()