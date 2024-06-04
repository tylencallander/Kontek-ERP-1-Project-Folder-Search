import os
import json
import openpyxl

# Added another key for Project Numbers not found in spreadsheet

basepath = 'P:/KONTEK/CUSTOMER'
projects = {}
errors = {
    "PROJECTNUMBERNOTINSPREADSHEET": [],
    "PROJECTNUMBERFOLDERNOTFOUND": []
}

# Extracts project numbers, including those with suffixes from the Excel file.

def extract_project_numbers_from_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        project_numbers = set()
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
            cell_value = str(row[0]).strip().upper() if row[0] else ''
            if cell_value.startswith('K') and len(cell_value) == 8 and cell_value[1:8].isdigit():
                project_numbers.add(cell_value)
                print(f"Extracted project number: {cell_value} from Excel sheet")
        return project_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set()
    
# Checks through the network folders to find project folders and logs them if they match the project numbers found in the Excel file.

def check_project_folder(base_path, excel_project_numbers):
    network_projects = set()
    try:
        for root, dirs, files in os.walk(base_path, topdown=True):
            for folder in dirs:
                full_path = os.path.join(root, folder)
                if folder.startswith('K') and len(folder) >= 8 and folder[1:8].isdigit():
                    project_number = folder[:8]
                    projects[project_number] = {
                        "projectnumber": project_number,
                        "projectfullpath": full_path,
                        "projectpath": full_path.split("\\")
                    }
                    network_projects.add(project_number)
                    print(f"Found and logged project: {project_number} at {full_path}")

        extra_projects = network_projects.difference(excel_project_numbers)
        errors["PROJECTNUMBERNOTINSPREADSHEET"].extend(extra_projects)
    except Exception as e:
        print(f"Error checking project folder: {e}")

# Finds unmatched projects in both the Excel file and network and prints them out for the user to fix

def find_unmatched_projects(excel_project_numbers, network_projects):
    missing_projects = excel_project_numbers.difference(network_projects)
    if missing_projects:
        errors["PROJECTNUMBERFOLDERNOTFOUND"].extend(missing_projects)
        for mp in missing_projects:
            print(f"Missing project number: {mp} not found in directories")

def main():
    print("\nParsing all Files in KONTEK's Network...\n")
    excel_file_path = "P:/KONTEK/KONTEK PROJECT JOB NUMBERS.xlsx"
    excel_project_numbers = extract_project_numbers_from_excel(excel_file_path)
    check_project_folder(basepath, excel_project_numbers)
    find_unmatched_projects(excel_project_numbers, set(projects.keys()))

# Creating projects.json and errors.json files to store parsed data

    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} found projects to projects.json")
    print(f"Logged {len(errors['PROJECTNUMBERNOTINSPREADSHEET'])} projects not in spreadsheet to errors.json")
    print(f"Logged {len(errors['PROJECTNUMBERFOLDERNOTFOUND'])} missing projects to errors.json")

if __name__ == "__main__":
    main()