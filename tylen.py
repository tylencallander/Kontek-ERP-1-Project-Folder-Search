import os
import json
import openpyxl

PROJECTS_ONLY = False

basepath = "P:/KONTEK/CUSTOMER"
projects = {}
errors = {}

def check_project_folder(letter, customerFolder, projectFolder):
    base_folder = os.path.join(basepath, letter, customerFolder, projectFolder)
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
    for root, dirs, files in os.walk(basepath):
        for customerFolder in dirs:
            customer_path = os.path.join(root, customerFolder)
            for projectFolder in os.listdir(customer_path):
                check_project_folder(os.path.basename(root), customerFolder, projectFolder)

def parse_project_numbers_xlsx():
    project_numbers = set()
    xlsx_path = "P:/KONTEK/KONTEK PROJECT JOB NUMBERS.xlsx"
    if os.path.exists(xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
            project_number = row[0].value
            if project_number.startswith("K"):
                project_numbers.add(project_number)
    return project_numbers

def save_json():
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4, sort_keys=True)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4, sort_keys=True)

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} projects to project.json")

    project_numbers_from_xlsx = parse_project_numbers_xlsx()
    projects_not_found = project_numbers_from_xlsx - set(projects.keys())
    if projects_not_found:
        errors["PROJECTNUMBERSFOLDERNOTFOUND"] = list(projects_not_found)
        print(f"Logged {len(projects_not_found)} project numbers not found in folders to errors.json")

def main():
    parse_projects()
    save_json()
    if PROJECTS_ONLY:
        exit()
    print("\nExiting Now...")

if __name__ == "__main__":
    main()

