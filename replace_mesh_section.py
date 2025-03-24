import re

# === CONFIG ===
source_path = r"C:\Users\Ryan\Downloads\Dogbone_7_unzipped\3D\3dmodel.model"
target_path = r"C:\Users\Ryan\Downloads\NumberedBonesPre_unzipped\3D\3dmodel.model"
output_path = r"C:\Users\Ryan\Downloads\NumberedBonesPre_unzipped\3D\3dmodel.model"
object_id_to_replace = "1"  # Update this if needed

# === FUNCTION ===
def replace_mesh_textually(source_path, target_path, output_path, object_id):
    with open(source_path, 'r', encoding='utf-8') as f:
        source_text = f.read()
    with open(target_path, 'r', encoding='utf-8') as f:
        target_text = f.read()

    # Extract <mesh>...</mesh> from source
    mesh_match = re.search(r"<mesh>.*?</mesh>", source_text, re.DOTALL)
    if not mesh_match:
        raise ValueError("Could not find <mesh> section in source file.")
    new_mesh = mesh_match.group(0)

    # Find <object id="X" type="model"> and its <mesh> section in target
    object_pattern = fr'(<object id="{object_id}" type="model">.*?)(<mesh>.*?</mesh>)(.*?</object>)'
    match = re.search(object_pattern, target_text, re.DOTALL)
    if not match:
        raise ValueError(f"Could not find object id={object_id} with a <mesh> in target file.")

    before_mesh = match.group(1)
    after_mesh = match.group(3)
    replaced_object = before_mesh + new_mesh + after_mesh

    # Replace the entire object block in the target text
    new_text = re.sub(object_pattern, re.escape(replaced_object), target_text, flags=re.DOTALL)
    new_text = new_text.replace(re.escape(replaced_object), replaced_object)  # Fix escaped text

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_text)

    print(f"✅ Mesh in object id={object_id} replaced and written to {output_path}")

# === RUN ===
replace_mesh_textually(source_path, target_path, output_path, object_id_to_replace)
