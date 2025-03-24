import subprocess
import os

# === CONFIG ===
PRUSA_SLICER_CLI = r"C:\Program Files\Prusa3D\PrusaSlicer\prusa-slicer-console.exe"
INPUT_3MF = r"C:\Users\Ryan\Downloads\PostTest_3.3mf"
OUTPUT_GCODE = r"C:\Users\Ryan\Downloads\G_PostTest_3.gcode"

def export_gcode(cli, input_file, output_file):
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
        print(f"G-code exported to: {output_file}")

# === RUN ===
export_gcode(PRUSA_SLICER_CLI, INPUT_3MF, OUTPUT_GCODE)
