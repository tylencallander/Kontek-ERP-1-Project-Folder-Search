import os
import json
import openpyxl

PROJECTS_ONLY = False

basepath = 'P:/KONTEK/CUSTOMER'
projects = {}
errors = {}

def check_project_folder(base_path):
    """Recursively checks each folder starting from the base path to find valid project folders."""
    for root, dirs, files in os.walk(base_path, topdown=True):
        for folder in dirs:
            full_path = os.path.join(root, folder)
            # Ignore if it's not a directory or doesn't start with 'K'
            if not folder.startswith('K') or not os.path.isdir(full_path):
                continue

            project_number = folder.replace("-", "").replace(" ", "")
            if not project_number[1:8].isnumeric():
                errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(full_path)
                continue

            # Skipping folders with 'OBSOLETE' in the name
            if "OBSOLETE" in folder.upper():
                errors.setdefault("OBSOLETEPROJECTFOLDER", []).append(full_path)
                continue

            # Base Project Info Object
            project_info = {
                "projectnumber": 'K' + project_number[1:8],
                "projectfullpath": full_path,
                "projectpath": full_path.split("\\")
            }
            projects[project_info['projectnumber']] = project_info
            print(f"Found project: {project_info['projectnumber']} at {full_path}")

# Main function to orchestrate the directory scanning and JSON saving
def main():
    print("********************")
    print("Parsing All Projects")
    print("********************")
    check_project_folder(basepath)

    # Saving the results to JSON files
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!")
    print(f"Logged {len(projects)} projects to projects.json")
    print(f"Logged {len(errors)} different types of errors to errors.json")

    if PROJECTS_ONLY:
        exit()

if __name__ == "__main__":
    main()
