import os
import ast
import xlwt

# Function to extract 'depends' and 'name' from __manifest__.py
def extract_module_info(manifest_path):
    with open(manifest_path, 'r') as manifest_file:
        manifest_content = manifest_file.read()
        manifest_data = ast.literal_eval(manifest_content)
        depends = manifest_data.get('depends', [])
        return depends

# Function to detect Odoo modules and extract info
def detect_modules_and_generate_excel(root_dir, output_file):
    # Create Excel workbook and worksheet
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Modules")

    # Write headers
    sheet.write(0, 0, 'Module Name')
    sheet.write(0, 1, 'Depends')

    row = 1  # Start writing from the second row

    # Traverse directory
    for module in os.listdir(root_dir):
        module_dir = os.path.join(root_dir, module)
        manifest_path = os.path.join(module_dir, '__manifest__.py')

        # Check if manifest file exists
        if os.path.isfile(manifest_path):
            depends = extract_module_info(manifest_path)

            # Write module name and depends to Excel sheet
            sheet.write(row, 0, module)
            sheet.write(row, 1, ",".join(depends))
            row += 1

    # Save workbook
    workbook.save(output_file)
    print(f"Excel file saved as {output_file}")

if __name__ == "__main__":
    # Ask the user for the root directory
    root_directory = input("Enter the path to the directory containing Odoo modules: ")

    # Get the path of the script and use it to save the output Excel file
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_excel_file = os.path.join(script_directory, 'odoo_modules.xls')

    # Run the detection and excel generation
    detect_modules_and_generate_excel(root_directory, output_excel_file)
