import os
import win32com.client as win32

def add_vba_macro_and_run(excel_path, image_folder, vba_file_name):
    excel = None
    try:
        # Open Excel
        #excel = win32.Dispatch("Excel.Application")
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True  # Make Excel visible if you want to see the process

        # Open the workbook
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))

        ws = wb.Sheets("Images")
        ws.Protect("password")
        ws.Unprotect("password")

        # Add VBA macro to the workbook
        excel_module = wb.VBProject.VBComponents.Add(1)  # 1 = Module

        # Read the VBA code from the file
        with open(vba_file_name, 'r') as vba_file:
            vba_code = vba_file.read()

        # Replace the placeholder with the actual image folder
        vba_code = vba_code.replace("<image_folder>", image_folder)

        # Add the VBA code to the Excel module
        excel_module.CodeModule.AddFromString(vba_code)

        # Run the VBA macro
        excel.Application.Run("InsertImagesBasedOnSKU")

        # Save and close the workbook
        wb.Save()
        wb.Close()
    except Exception as e:
        print(f"An error occurred while adding the VBA macro or running it: {e}")
    finally:
        if excel:
            excel.Quit()




