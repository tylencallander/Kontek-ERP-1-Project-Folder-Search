import os
import json

PROJECTS_ONLY = False

basepath = 'P:/KONTEK/CUSTOMER'
letters = ['#', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

projects = {}
errors = {}

def check_project_folder(letter, customerFolder, projectFolder):
    base_folder = f"{basepath}/{letter}/{customerFolder}/{projectFolder}"
    if not os.path.isdir(base_folder):
        return
    
    projectnumber = projectFolder.replace("-", "").replace(" ", "")
    if not projectnumber.startswith('K'):  # Verify project number starts with 'K'
        return  # Ignore folder if it does not start with 'K'
    
    # Extract the numeric part following the initial 'K'
    if not projectnumber[1:8].isnumeric():
        errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(base_folder)
        return

    # Use only the first 8 characters, K + 7 digits
    project_number_formatted = 'K' + projectnumber[1:8]
    projects[project_number_formatted] = {
        "projectnumber": project_number_formatted,
        "projectfullpath": base_folder,
        "projectpath": ["P:", "KONTEK", "CUSTOMER", letter, customerFolder, projectFolder]
    }

def parse_projects():
    for letter in letters:
        path = f"{basepath}/{letter}"
        try:
            contents = os.listdir(path)
        except FileNotFoundError:
            continue  

        for customerFolder in contents:
            customer_path = f"{path}/{customerFolder}"
            if os.path.isdir(customer_path):
                projectsList = os.listdir(customer_path)
                for projectFolder in projectsList:
                    check_project_folder(letter, customerFolder, projectFolder)

def save_json():
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
