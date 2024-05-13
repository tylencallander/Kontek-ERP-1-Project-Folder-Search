import os
import json
import pandas as pd

PROJECTS_ONLY = False

basepath = "P:/KONTEK/CUSTOMER"
projects = {}
errors = {}

print("\nParsing all Files in KONTEK's Network...")

def check_project_folder(folder_path):
    for folder in os.listdir(folder_path):
        folder_full_path = os.path.join(folder_path, folder)
        if os.path.isdir(folder_full_path):
            project_number = folder.replace("-", "").replace(" ", "")
            if not project_number.startswith("K"):
                continue
            if not project_number[1:8].isnumeric():
                errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(folder_full_path)
                continue
            project_number_formatted = 'K' + project_number[1:8]
            projects[project_number_formatted] = {
                "projectnumber": project_number_formatted,
                "projectfullpath": folder_full_path,
                "projectpath": folder_full_path.split("\\")
            }

def parse_projects():
    for root, dirs, files in os.walk(basepath):
        for d in dirs:
            check_project_folder(os.path.join(root, d))

def check_missing_projects():
    df = pd.read_excel("P:/KONTEK/KONTEK PROJECT JOB NUMBERS.xlsx")
    for project_number in df["Project Number"]:
        project_number = str(project_number).strip().upper()
        if project_number.startswith("K"):
            if project_number not in projects:
                errors.setdefault("PROJECTNUMBERSFOLDERNOTFOUND", []).append(project_number)

def save_json():
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4, sort_keys=True)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4, sort_keys=True)
    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} projects to project.json")
    print(f"Logged {len(errors['PROJECTNUMBERNOTNUMERIC'])} non-numeric projects errors to errors.json")
    print(f"Logged {len(errors['PROJECTNUMBERSFOLDERNOTFOUND'])} projects with no folder found to errors.json")

def main():
    parse_projects()
    check_missing_projects()
    save_json()
    if PROJECTS_ONLY:
        exit()
    print("\nExiting Now...")

if __name__ == "__main__":
    main()
