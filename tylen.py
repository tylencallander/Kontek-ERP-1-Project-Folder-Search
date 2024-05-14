import os
import json
import openpyxl

PROJECTS_ONLY = False

basepath = 'P:/KONTEK/CUSTOMER'
projects = {}
errors = {}

def extract_project_numbers_from_excel(file_path):
    """Extract project numbers from the provided Excel file that start with 'K'."""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    # Extract and filter project numbers that start with 'K'
    return {str(cell[0]).strip().upper() for cell in ws.iter_rows(min_row=2, max_col=1, values_only=True) if cell[0] and str(cell[0]).strip().upper().startswith('K')}

def check_project_folder(base_path):
    """Recursively checks each folder starting from the base path to find valid project folders that start with 'K'."""
    for root, dirs, files in os.walk(base_path, topdown=True):
        for folder in dirs:
            full_path = os.path.join(root, folder)
            if not folder.startswith('K') or not os.path.isdir(full_path):
                continue

            project_number = 'K' + folder[1:8] if folder[1:8].isdigit() else None
            if project_number and len(project_number) == 8:  # Confirm project number length
                projects[project_number] = {
                    "projectnumber": project_number,
                    "projectfullpath": full_path,
                    "projectpath": full_path.split("\\")
                }
                print(f"Found project: {project_number} at {full_path}")

def find_missing_projects(excel_project_numbers):
    """Log project numbers that are in the Excel but not found in the directories."""
    found_projects = set(projects.keys())
    missing_projects = excel_project_numbers.difference(found_projects)
    if missing_projects:
        errors["PROJECTNUMBERSNOTFOUND"] = list(missing_projects)
        for mp in missing_projects:
            print(f"Project number in Excel not found in network: {mp}")

def save_json():
    """Save found projects and errors to JSON files."""
    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print(f"\nParsing Complete!\n")
    print(f"Logged {len(projects)} projects to projects.json")
    print(f"Logged {len(errors.get('PROJECTNUMBERSNOTFOUND', []))} projects not found in network to errors.json")

def main():
    print("\nParsing all Files in KONTEK's Network...")

    excel_project_numbers = extract_project_numbers_from_excel("P:/KONTEK/KONTEK PROJECT JOB NUMBERS.xlsx")
    check_project_folder(basepath)
    find_missing_projects(excel_project_numbers)

    save_json()

    if PROJECTS_ONLY:
        print("\nExiting Now...")
        exit()

if __name__ == "__main__":
    main()
