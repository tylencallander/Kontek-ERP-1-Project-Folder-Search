# Was going to use Pandas or Openpyxl to work with the Excel file but decided not to instead.

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
    if not projectnumber.startswith('K'):  
        return  
    
    if not projectnumber[1:8].isnumeric():
        errors.setdefault("PROJECTNUMBERNOTNUMERIC", []).append(base_folder)
        return

    # Check if project number is 7 characters long plus the K prefix, NOT INCLUDING SUFFIXES LIKE -E, -R etc..

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

# Save parsed files to projects.json if located, then errors.json if the value is not numeric (8>x<8 characters)

def save_json():
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4, sort_keys=True)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4, sort_keys=True)
        
# Added some print statements so I can debug and confirm that all data from the spreadsheet path has been parsed successfully or populated into the error file
# They can be omitted if unecessary, doesn't matter to me

def main():
    parse_projects()
    save_json()
    if PROJECTS_ONLY:
        exit()

if __name__ == "__main__":
    main()

# Missing all of the projects that dont begin with K prefix, ill come back to this later.
# Missing all of the projects that dont end with the alphabetical suffixes, ill come back to this later.