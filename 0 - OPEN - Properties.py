import os
import subprocess

scripts_path = "C:\\Users\\johnj\\Desktop\\CoastRev\\Reporting\\Scripts"
scripts_to_run = []

# Look for .py files that start with "0 - R" and are not in subfolders
for filename in os.listdir(scripts_path):
    if filename.startswith("0 - R") and filename.endswith(".py"):
        full_path = os.path.join(scripts_path, filename)
        if os.path.isfile(full_path):
            scripts_to_run.append(full_path)

failed_scripts = []

for script in scripts_to_run:
    filename = os.path.basename(script)
    print(f"Running script: {filename}")
    try:
        subprocess.run(["python", script], check=True)
        print(f"Completed script: {filename}\n")
    except subprocess.CalledProcessError:
        print(f"FAILED to run script: {filename}\n")
        failed_scripts.append(filename)

if failed_scripts:
    print("These scripts did not complete successfully:")
    for failed in failed_scripts:
        print(f"- {failed}")
else:
    print("All scripts completed successfully.")
