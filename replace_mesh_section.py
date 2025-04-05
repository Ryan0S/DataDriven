import re
import shutil
import os
import xml.etree.ElementTree as ET
import zipfile
import tempfile
import json
import subprocess
from typing import List, Dict, Optional
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime



# === API CONFIGURATION ===
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'  # Your OAuth Client ID JSON
TOKEN_FILE = 'token_drive.pickle'      # Will be created automatically
DRIVE_FOLDER_ID = '1rXqywoBKWUcdExU4jpIWjZ4mwbOlmHZF'  # Target folder inside shared drive
SHARED_DRIVE_ID = '0ADuXLLjfuoALUk9PVA'   # The shared drive ID (starts with '0A...')

SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SPREADSHEET_ID = '1Jnc-Lyju3zLz7i5Xaq7G5_8k_c0LYy8FLVQSFWNW1Ds'
TOKEN_SHEETS_FILE = 'token_sheets.pickle'
SHEET_NAME = 'objects_json'


# === GCODE and 3MF CONFIG ===
CONFIG = {
    "template_3mf_file": r"C:\Users\Ryan\Downloads\template_11_dogbones.3mf",
    "output_3mf_folder_base": r"C:\Users\Ryan\Downloads\11_dogbones",
    "output_3mf_file_base": r"C:\Users\Ryan\Downloads\11_",
    "output_gcode_base": r"C:\Users\Ryan\Downloads\11_",  # Base name for G-code files
    "max_objects_per_file": 11,  # Maximum number of objects per output file
    "objects_json_file": r"C:\Users\Ryan\Desktop\Tidy\DataDriven\objects.json",
    "prusa_slicer_cli": r"C:\Program Files\Prusa3D\PrusaSlicer\prusa-slicer-console.exe"  # Path to PrusaSlicer CLI
}

# Dynamic source file path generator
SOURCE_FILE_BASE_DIR = r"C:\Users\Ryan\Downloads\DogboneExports"

def get_sheets_service():
    """Authenticate and return Sheets service."""
    creds = None
    if os.path.exists(TOKEN_SHEETS_FILE):
        with open(TOKEN_SHEETS_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_SHEETS_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build('sheets', 'v4', credentials=creds)

def fetch_objects_json():
    """Fetch and parse the objects_json content from the sheet."""
    service = get_sheets_service()
    range_name = f"{SHEET_NAME}!A1"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', [])

    if not values:
        raise ValueError("❌ No data found in objects_json sheet.")

    # The JSON is in A1
    json_text = values[0][0]
    try:
        objects = json.loads(json_text)
        print(f"✅ Fetched {len(objects)} objects from sheet.")
        return objects
    except json.JSONDecodeError:
        raise ValueError("❌ Failed to parse JSON content from sheet.")


def get_drive_service():
    """Authenticate and return the Drive service."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build('drive', 'v3', credentials=creds)


def upload_to_shared_drive(file_path, folder_id, shared_drive_id):
    """Upload file to a Shared Drive folder."""
    service = get_drive_service()
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path, resumable=True)
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id',
        supportsAllDrives=True
    ).execute()
    print(f'✅ Uploaded to Shared Drive: {file_path} (File ID: {file.get("id")})')

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
        for obj in objects:
            if not all(key in obj for key in ["specimen_id", "fill_density", "fill_pattern"]):
                raise ValueError(f"❌ Invalid object in JSON: {obj}. Must include specimen_id, fill_density, and fill_pattern.")
        return objects
    except FileNotFoundError:
        raise FileNotFoundError(f"❌ JSON file not found: {json_file}")
    except json.JSONDecodeError:
        raise ValueError(f"❌ Invalid JSON format in file: {json_file}")

def export_gcode(cli: str, input_file: str, output_file: str) -> None:
    """Slice a .3mf file and export it as G-code using PrusaSlicer CLI."""
    cmd = [
        cli,
        "--printer-technology", "FFF",
        "--slice",
        "--export-gcode",
        "-o", output_file,
        input_file
    ]

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print("❌ Error during slicing:")
        print("STDOUT:\n", result.stdout)
        print("STDERR:\n", result.stderr)
    else:
        print(f"✅ G-code exported to: {output_file}")

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
            solid_infill_every_layers = obj_config["solid_infill_every_layers"]
            perimeters = obj_config["perimeters"]


            source_path = os.path.join(source_paths[specimen_id], "3D", "3dmodel.model")
            source_text = read_file(source_path)
            triangle_count = self.count_triangles(source_text)

            obj = root.find(f".//object[@id='{object_id}']")
            if obj is None:
                raise ValueError(f"❌ Could not find object id={object_id} in config file.")

            self._update_metadata(obj, "fill_density", fill_density)
            self._update_metadata(obj, "fill_pattern", fill_pattern)
            self._update_metadata(obj, "solid_infill_every_layers", solid_infill_every_layers)
            self._update_metadata(obj, "perimeters", perimeters)
            
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
    objects = fetch_objects_json()

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
                # Get specimen ID range
                first_id = group_objects[0]["specimen_id"]
                last_id = group_objects[-1]["specimen_id"]

                # Get timestamp
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")

                # Build filenames
                base_name = f"{first_id}-{last_id}_{timestamp}"
                output_folder = f"{CONFIG['output_3mf_folder_base']}_{file_idx + 1}"
                output_3mf_file = f"{CONFIG['output_3mf_file_base']}{base_name}.3mf"
                output_gcode_file = f"{CONFIG['output_gcode_base']}{base_name}.gcode"


                # Process and generate .3mf file
                processor = ModelProcessor(template_folder, output_folder)
                processor.process(group_objects, source_paths)
                rezip_3mf(output_folder, output_3mf_file)

                # Slice .3mf file to G-code
                export_gcode(CONFIG["prusa_slicer_cli"], output_3mf_file, output_gcode_file)

                upload_to_shared_drive(output_3mf_file, DRIVE_FOLDER_ID, SHARED_DRIVE_ID)
                upload_to_shared_drive(output_gcode_file, DRIVE_FOLDER_ID, SHARED_DRIVE_ID)


        finally:
            # Clean up temporary directories
            for temp_dir in temp_dirs.values():
                temp_dir.cleanup()

if __name__ == "__main__":
    main()