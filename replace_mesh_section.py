import re
import shutil
import os
import xml.etree.ElementTree as ET
import zipfile
import tempfile

# === CONFIG ===
source_3mf_file = r"C:\Users\Ryan\Downloads\Dogbone_7.3mf"
template_3mf_file = r"C:\Users\Ryan\Downloads\NumberedBonesPre.3mf"
output_3mf_folder = r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified"
output_3mf_file = r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified.3mf"
object_id_to_replace = "1"
new_fill_density = "70%"    # New fill density value
new_fill_pattern = "grid"   # New fill pattern value

# === FUNCTIONS ===
def extract_3mf(archive_path, extract_to):
    """Extract a .3mf file to a specified directory."""
    with zipfile.ZipFile(archive_path, 'r') as z:
        z.extractall(extract_to)
    print(f"Extracted {archive_path} to {extract_to}")
    return extract_to

def replace_mesh_in_3mf_folder(source_folder, template_folder, output_folder, object_id, fill_density, fill_pattern):
    source_model_path = os.path.join(source_folder, "3D", "3dmodel.model")
    template_model_path = os.path.join(template_folder, "3D", "3dmodel.model")
    
    # Duplicate template folder into output
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    shutil.copytree(template_folder, output_folder)
    output_model_path = os.path.join(output_folder, "3D", "3dmodel.model")
    output_config_path = os.path.join(output_folder, "Metadata", "Slic3r_PE_model.config")

    # Read 3dmodel.model files
    with open(source_model_path, 'r', encoding='utf-8') as f:
        source_text = f.read()
    with open(template_model_path, 'r', encoding='utf-8') as f:
        target_text = f.read()

    # Count number of <triangle ... /> entries in the source
    triangle_count = len(re.findall(r"<triangle\b", source_text)) - 1

    # Extract <mesh>...</mesh> from source
    mesh_match = re.search(r"<mesh>.*?</mesh>", source_text, re.DOTALL)
    if not mesh_match:
        raise ValueError("❌ Could not find <mesh> section in source file.")
    new_mesh = mesh_match.group(0)

    # Find <object id="X" type="model"> and its <mesh> section in target
    object_pattern = fr'(<object id="{object_id}" type="model">.*?)(<mesh>.*?</mesh>)(.*?</object>)'
    match = re.search(object_pattern, target_text, re.DOTALL)
    if not match:
        raise ValueError(f"❌ Could not find object id={object_id} with a <mesh> in target file.")

    before_mesh = match.group(1)
    after_mesh = match.group(3)
    replaced_object = before_mesh + new_mesh + after_mesh

    # Replace the mesh section in target text
    new_text = re.sub(object_pattern, re.escape(replaced_object), target_text, flags=re.DOTALL)
    new_text = new_text.replace(re.escape(replaced_object), replaced_object)

    # Write to output model file
    with open(output_model_path, 'w', encoding='utf-8') as f:
        f.write(new_text)

    # Update the config file
    tree = ET.parse(output_config_path)
    root = tree.getroot()
    
    # Find the object with matching id and update its properties
    for obj in root.findall(f".//object[@id='{object_id}']"):
        # Update fill_density
        fill_density_elem = obj.find(".//metadata[@key='fill_density']")
        if fill_density_elem is not None:
            fill_density_elem.set("value", fill_density)
        else:
            ET.SubElement(obj, "metadata", {"type": "object", "key": "fill_density", "value": fill_density})

        # Update fill_pattern
        fill_pattern_elem = obj.find(".//metadata[@key='fill_pattern']")
        if fill_pattern_elem is not None:
            fill_pattern_elem.set("value", fill_pattern)
        else:
            ET.SubElement(obj, "metadata", {"type": "object", "key": "fill_pattern", "value": fill_pattern})

        # Update volume's lastid
        volume = obj.find("volume")
        if volume is not None:
            volume.set("lastid", str(triangle_count))
        else:
            raise ValueError(f"❌ Could not find volume for object id={object_id} in config file.")
        break
    else:
        raise ValueError(f"❌ Could not find object id={object_id} in config file.")

    # Write updated config file
    tree.write(output_config_path, encoding="utf-8", xml_declaration=True)

    print(f"✅ Mesh replaced, config updated (fill_density={fill_density}, fill_pattern={fill_pattern}, lastid={triangle_count}) "
          f"for object id={object_id} in:\n{output_model_path}")
    
    return output_folder

def rezip_3mf(folder_path, output_path):
    """Repack the contents of a folder into a valid .3mf (zip) file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder_path):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, folder_path)
                z.write(abs_path, rel_path)
    print(f"✅ Repacked into: {output_path}")

# === RUN ===
def main():
    # Create temporary directories for extraction
    with tempfile.TemporaryDirectory() as temp_source_dir, tempfile.TemporaryDirectory() as temp_template_dir:
        # Step 1: Extract source and template .3mf files
        source_folder = extract_3mf(source_3mf_file, temp_source_dir)
        template_folder = extract_3mf(template_3mf_file, temp_template_dir)

        # Step 2: Modify the 3MF folder contents
        modified_folder = replace_mesh_in_3mf_folder(
            source_folder,
            template_folder,
            output_3mf_folder,
            object_id_to_replace,
            new_fill_density,
            new_fill_pattern
        )
    
        # Step 3: Repack the modified folder into a .3mf file
        rezip_3mf(modified_folder, output_3mf_file)

    # Temporary directories are automatically cleaned up when exiting the 'with' block

if __name__ == "__main__":
    main()