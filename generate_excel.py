import os
import subprocess
import win32com.client as win32
from add_vba_macro_to_run import add_vba_macro_and_run

def kill_excel_processes():
    try:
        # List all processes
        result = subprocess.run(['tasklist'], stdout=subprocess.PIPE)
        processes = result.stdout.decode('utf-8')

        # Check if Excel is running
        if 'EXCEL.EXE' in processes:
            # Kill all Excel processes
            os.system("taskkill /f /im EXCEL.EXE")
            print("Killed all running Excel processes.")
        else:
            print("No Excel processes found.")
    except Exception as e:
        print(f"An error occurred while killing Excel processes: {e}")

def generate_excel_with_skus(image_folder, output_excel):
    excel = None
    try:
        # Start Excel application
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Keep Excel hidden

        # Add a new workbook
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Images"

        # Add column headers
        ws.Cells(1, 1).Value = "SKU"
        ws.Cells(1, 2).Value = "Image"
        ws.Cells(1, 3).Value = "Image Name"

        # Set column widths for better visibility
        ws.Columns("A").ColumnWidth = 15
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 15

        # Iterate through the images in the folder
        row = 2
        for filename in os.listdir(image_folder):
            if filename.endswith(".png"):
                # Extract the SKU from the filename
                sku = filename.split('.')[0]

                # Add the SKU to the Excel file
                ws.Cells(row, 1).Value = sku
                row += 1

        # Save the workbook
        wb.SaveAs(os.path.abspath(output_excel))
        wb.Close()
        print(f"Excel file with SKUs created: {output_excel}")

    except Exception as e:
        print(f"An error occurred while generating the Excel file: {e}")
    finally:
        if excel:
            excel.Quit()


if __name__ == '__main__':
    try:
        kill_excel_processes()
        image_folder = os.path.abspath('.\\test_images\\')
        output_excel = 'test.xlsx'
        
        # Generate the Excel file with SKU numbers
        generate_excel_with_skus(image_folder, output_excel)
        
        # Add the VBA macro and run it to insert images and write image names
        vba_macro = ".\\vba\\insert_images_macro.vba"
        add_vba_macro_and_run(output_excel, image_folder, vba_macro)
        print("VBA macro added and executed to insert images and write image names.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        kill_excel_processes()
