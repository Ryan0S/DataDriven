import re
import shutil
import os
import xml.etree.ElementTree as ET
import zipfile

# === CONFIG ===
source_3mf_folder = r"C:\Users\Ryan\Downloads\Dogbone_7_unzipped"
template_3mf_folder = r"C:\Users\Ryan\Downloads\NumberedBonesPre_unzipped"
output_3mf_folder = r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified"
output_3mf_file = r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified.3mf"
object_id_to_replace = "1"

# === FUNCTIONS ===
def replace_mesh_in_3mf_folder(source_folder, template_folder, output_folder, object_id):
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

    # Update the config file with triangle count
    tree = ET.parse(output_config_path)
    root = tree.getroot()
    
    # Find the object with matching id and update its volume's lastid
    for obj in root.findall(f".//object[@id='{object_id}']"):
        volume = obj.find("volume")
        if volume is not None:
            volume.set("lastid", str(triangle_count))
            break
    else:
        raise ValueError(f"❌ Could not find object id={object_id} in config file.")

    # Write updated config file
    tree.write(output_config_path, encoding="utf-8", xml_declaration=True)

    print(f"✅ Mesh replaced and config updated for object id={object_id} in duplicated 3MF folder at:\n{output_model_path}")
    
    return output_folder  # Return the modified folder path for zipping

def rezip_3mf(folder_path, output_path):
    """
    Repack the contents of a folder into a valid .3mf (zip) file.
    - folder_path: path to folder containing 3D/, Metadata/, etc.
    - output_path: desired output .3mf path
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder_path):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, folder_path)
                z.write(abs_path, rel_path)
    print(f"✅ Repacked into: {output_path}")

# === RUN ===
def main():
    # Step 1: Modify the 3MF folder contents
    modified_folder = replace_mesh_in_3mf_folder(
        source_3mf_folder,
        template_3mf_folder,
        output_3mf_folder,
        object_id_to_replace
    )
    
    # Step 2: Repack the modified folder into a .3mf file
    rezip_3mf(modified_folder, output_3mf_file)

if __name__ == "__main__":
    main()