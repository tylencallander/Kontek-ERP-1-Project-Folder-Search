import os
import json

PROJECTS_ONLY = False

basepath = "P:/KONTEK/CUSTOMER"
letters = ["#", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

projects = {}
errors = {}

# Added print statements so I can debug while I was working on this, but they can be omitted.

print("\nParsing all Files in KONTEK's Network...")

def check_project_folder(letter, customerFolder, projectFolder):
    base_folder = f"{basepath}/{letter}/{customerFolder}/{projectFolder}"
    if not os.path.isdir(base_folder):
        return
    
    projectnumber = projectFolder.replace("-", "").replace(" ", "")
    if not projectnumber.startswith("K"):  
        return  
    
    if not projectnumber[1:8].isnumeric():
        errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(base_folder)
        return

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
    nested_paths = sum(len(value) for value in errors.values())

# More print statements for debugging.

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} projects to project.json")
    nested_paths = sum(len(value) for value in errors.values())
    print(f"Logged {nested_paths} non-numeric projects errors to errors.json")

def main():
    parse_projects()
    save_json()
    if PROJECTS_ONLY:
        exit()
    print("\nExiting Now...")

if __name__ == "__main__":
    main()
    