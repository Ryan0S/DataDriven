import re
import shutil
import os

# === CONFIG ===
source_3mf_folder = r"C:\Users\Ryan\Downloads\Dogbone_7_unzipped"
template_3mf_folder = r"C:\Users\Ryan\Downloads\NumberedBonesPre_unzipped"
output_3mf_folder = r"C:\Users\Ryan\Downloads\NumberedBones_7_Modified"
object_id_to_replace = "1"

# === FUNCTION ===
def replace_mesh_in_3mf_folder(source_folder, template_folder, output_folder, object_id):
    source_model_path = os.path.join(source_folder, "3D", "3dmodel.model")
    template_model_path = os.path.join(template_folder, "3D", "3dmodel.model")

    # Duplicate template folder into output
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    shutil.copytree(template_folder, output_folder)
    output_model_path = os.path.join(output_folder, "3D", "3dmodel.model")

    # Read files
    with open(source_model_path, 'r', encoding='utf-8') as f:
        source_text = f.read()
    with open(template_model_path, 'r', encoding='utf-8') as f:
        target_text = f.read()

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

    print(f"✅ Mesh replaced for object id={object_id} in duplicated 3MF folder at:\n{output_model_path}")

# === RUN ===
replace_mesh_in_3mf_folder(source_3mf_folder, template_3mf_folder, output_3mf_folder, object_id_to_replace)
