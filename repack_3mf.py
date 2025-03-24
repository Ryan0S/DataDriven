import os
import zipfile

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
    print(f"Repacked into: {output_path}")

# --- Example usage ---
if __name__ == "__main__":
    input_folder = r"C:\Users\Ryan\Downloads\NumberedBonesPre_unzipped"
    output_3mf = r"C:\Users\Ryan\Downloads\PostTest_3.3mf"
    rezip_3mf(input_folder, output_3mf)
