import os
import json


import os
import json
import xlrd
import openpyxl

PROJECTS_ONLY = False

basepath = 'P:/KONTEK/CUSTOMER'
letters = ['#', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

projects = {}
errors = {}

def check_project_folder(letter, customerFolder, projectFolder):
    """ Check the project folder and record errors and details. """
    base_folder = f"{basepath}/{letter}/{customerFolder}/{projectFolder}"
    if not os.path.isdir(base_folder):
        return  # Ignore non-directory items
    
    # Check for valid project number
    projectnumber = projectFolder.replace("-", "").replace(" ", "")
    if not projectnumber[1:8].isnumeric():
        errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(base_folder)
        return

    # Record project if valid
    projects[f"K{projectnumber[1:8]}"] = {
        "projectnumber": f"K{projectnumber[1:8]}",
        "projectfullpath": base_folder,
        "projectpath": ["P:", "KONTEK", "CUSTOMER", letter, customerFolder, projectFolder]
    }

def parse_projects():
    """ Parses all projects based on predefined letters and basepath. """
    for letter in letters:
        path = f"{basepath}/{letter}"
        try:
            contents = os.listdir(path)
        except FileNotFoundError:
            continue  # Skip if the directory does not exist

        for customerFolder in contents:
            customer_path = f"{path}/{customerFolder}"
            if os.path.isdir(customer_path):
                projectsList = os.listdir(customer_path)
                for projectFolder in projectsList:
                    check_project_folder(letter, customerFolder, projectFolder)

def save_json():
    """ Saves projects and errors to JSON files. """
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4, sort_keys=True)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4, sort_keys=True)

def main():
    parse_projects()
    save_json()
    if PROJECTS_ONLY:
        exit()

if __name__ == "__main__":
    main()
