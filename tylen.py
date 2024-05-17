import os
import json
import openpyxl

basepath = 'P:/KONTEK/CUSTOMER'
projects = {}
errors = {}

def extract_project_numbers_from_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        project_numbers = set()
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
            cell_value = str(row[0]).strip().upper() if row[0] else ''
            if cell_value.startswith('K') and cell_value[1:8].isdigit():
                project_numbers.add(cell_value)
                print(f"Extracted project number: {cell_value} from Excel sheet")
        return project_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set()

    
def check_project_folder(base_path):
    try:
        for root, dirs, files in os.walk(base_path, topdown=True):
            for folder in dirs:
                full_path = os.path.join(root, folder)
                # Split the folder name on spaces to isolate the project number part before any descriptions
                folder_parts = folder.split()
                for part in folder_parts:
                    # Check each part to find one that looks like a project number
                    if part.startswith('K') and len(part) >= 8 and part[1:8].isdigit():
                        # Now check for a suffix that's strictly letters and up to 3 characters long
                        if '-' in part:
                            project_base, suffix = part.split('-', 1)
                            if suffix.isalpha() and 1 <= len(suffix) <= 3:
                                project_number = project_base + '-' + suffix
                            else:
                                project_number = project_base
                        else:
                            project_number = part[:8]  # Take only the first 8 characters assuming they're formatted correctly
                        
                        # Log the project if it's correctly formatted
                        if project_number not in projects:
                            projects[project_number] = {
                                "projectnumber": project_number,
                                "projectfullpath": full_path,
                                "projectpath": full_path.split("\\")
                            }
                            print(f"Found and logged project: {project_number} at {full_path}")
                            break  # Stop after finding the first valid project number
    except Exception as e:
        print(f"Error checking project folder: {e}")



def find_unmatched_projects(excel_project_numbers):
    try:
        found_projects = set(projects.keys())
        missing_projects = excel_project_numbers.difference(found_projects)
        if missing_projects:
            errors["PROJECTNUMBERSFOLDERNOTFOUND"] = list(missing_projects)
            for mp in missing_projects:
                print(f"Missing project number: {mp} not found in directories")
    except Exception as e:
        print(f"Error finding unmatched projects: {e}")



def main():
    print("\nParsing all Files in KONTEK's Network...\n")
    excel_file_path = "P:/KONTEK/KONTEK PROJECT JOB NUMBERS.xlsx"
    excel_project_numbers = extract_project_numbers_from_excel(excel_file_path)
    check_project_folder(basepath)
    find_unmatched_projects(excel_project_numbers)

    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} found projects to projects.json")
    print(f"Logged {len(errors.get('PROJECTNUMBERSFOLDERNOTFOUND', []))} missing projects to errors.json")

if __name__ == "__main__":
    main()

