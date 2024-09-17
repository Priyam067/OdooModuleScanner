import os
import ast
import xlwt
import sys


# Function to extract relevant information from __manifest__.py
def extract_module_info(manifest_path):
    with open(manifest_path, 'r') as manifest_file:
        try:
            manifest_content = manifest_file.read()
            manifest_data = ast.literal_eval(manifest_content)
        except:
            print(f"Warning: {manifest_path} could not be parsed. may be any syntax issue!?")
            return None

        # Check if the parsed data is exactly one dictionary
        if not isinstance(manifest_data, dict):
            print(f"Warning: {manifest_path} does not Looks like odoo manifest.")
            return None
        return manifest_data


def create_header_style(workbook):
    style = xlwt.XFStyle()

    # Define font style
    font = xlwt.Font()
    font.bold = True
    font.colour_index = xlwt.Style.colour_map['white']
    style.font = font

    # Define background color
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_blue']
    style.pattern = pattern

    return style


# Function to create and return row style
def create_row_style(workbook):
    style = xlwt.XFStyle()

    # Define alignment and wrap text
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_LEFT
    alignment.vert = xlwt.Alignment.VERT_TOP
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    style.alignment = alignment

    return style


# Function to detect Odoo modules and extract info
def detect_modules_and_generate_excel(root_dir, output_file):
    # Create Excel workbook and worksheet
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Modules")

    # Create styles
    header_style = create_header_style(workbook)
    row_style = create_row_style(workbook)

    headers = [
        'Path to Module',
        'Module Name',
        'Version',
        'Price',
        'Author',
        'Dependent Modules',
        'Data Files',
        'Demo Files',
        'Assets',
        'Description',
        'Summary',
        'Additional Data'
    ]

    # Write initial headers with styling
    for idx, header in enumerate(headers):
        sheet.write(0, idx, header, header_style)

    # Traverse directory
    row = 1  # Start writing from the second row
    for module_dir, _, filenames in os.walk(root_dir):
        module = os.path.basename(module_dir)
        init_path = os.path.join(module_dir, '__init__.py')
        manifest_path = os.path.join(module_dir, '__manifest__.py')

        # Check if both files exist
        if os.path.isfile(init_path) and os.path.isfile(manifest_path):
            manifest_data = extract_module_info(manifest_path)
            if manifest_data:
                depends = manifest_data.get('depends', [])
                data = manifest_data.get('data', [])
                demo = manifest_data.get('demo', [])
                assets = manifest_data.get('assets', [])
                sheet.write(row, 0, module_dir, row_style)
                sheet.write(row, 1, module, row_style)
                sheet.write(row, 2, manifest_data.get('version'), row_style)
                sheet.write(row, 3, manifest_data.get('price'), row_style)
                sheet.write(row, 4, manifest_data.get('author'), row_style)
                sheet.write(row, 5, ",\n".join(depends) if isinstance(depends, list) else "", row_style)
                sheet.write(row, 6, ",\n".join(data) if isinstance(data, list) else "", row_style)
                sheet.write(row, 7, ",\n".join(demo) if isinstance(demo, list) else "", row_style)
                sheet.write(row, 8, ",\n".join(assets) if isinstance(assets, list) else "", row_style)
                sheet.write(row, 9, manifest_data.get('description'), row_style)
                sheet.write(row, 10, manifest_data.get('summary'), row_style)

                # additional data
                additional_data = [f"{key}: {value}" for key, value in manifest_data.items() if key not in [
                    'depends', 'description', 'summary', 'author', 'price', 'version', 'data', 'demo', 'assets']]
                sheet.write(row, 11, ",\n".join(additional_data), row_style)

                row += 1

    # Save workbook
    workbook.save(output_file)
    print(f"Excel file saved as {output_file}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python main.py /path/to/directory")
        sys.exit(1)

    root_directory = sys.argv[1]  # Get the directory path from the command line argument
    if not os.path.isdir(root_directory):
        print(f"Error: {root_directory} is not a valid directory.")
        sys.exit(1)

    # Get the path of the script and use it to save the output Excel file
    script_directory = os.path.dirname(os.path.abspath(__file__))
    output_excel_file = os.path.join(script_directory, 'odoo_modules_data.xls')

    # Run the detection and Excel generation
    detect_modules_and_generate_excel(root_directory, output_excel_file)
