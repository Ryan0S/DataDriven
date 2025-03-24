import subprocess
import os

# --- USER PARAMETERS ---
prusa_slicer_path = r"C:\Program Files\Prusa3D\PrusaSlicer\prusa-slicer-console.exe"
stl_file_path = r"C:\Users\Ryan\Downloads\Dogbone.stl"
output_dir = r"C:\Users\Ryan\Desktop\Tidy\DataDriven"
project_name = "dogbone_batch"

# XY-Z positions in mm for each model
positions = [
    (0, 0, 0),
    (40, 0, 0),
    (80, 0, 0),
    (0, 40, 0),
    (40, 40, 0),
    (80, 40, 0),
]

infill_values = [10, 20, 30, 40, 50, 60]  # One per dogbone

# Create output dir if needed
os.makedirs(output_dir, exist_ok=True)

# Create object load commands with per-object settings
stl_args = []
for i in range(6):
    x, y, z = positions[i]
    infill = infill_values[i]
    stl_args.extend([
        "--load-object", stl_file_path,
        "--object-settings",
        f"infill_density={infill}",
        "--place-object",
        f"{x},{y},{z}"
    ])

# File output paths
project_path = os.path.join(output_dir, f"{project_name}.3mf")
gcode_path = os.path.join(output_dir, f"{project_name}.gcode")

# Build full command
cmd = [
    prusa_slicer_path,
    "--export-3mf",
    "--export-gcode",
    "--output", gcode_path,
    "--save", project_path
] + stl_args


# Run the command
print("Running PrusaSlicer CLI...")
result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

# Check output
if result.returncode == 0:
    print(f"Success!\n3MF: {project_path}\nG-code: {gcode_path}")
else:
    print("Error running PrusaSlicer:")
    print(result.stderr)
