import xlwings as xw

def insert_picture_into_excel2(input_workbook_path, output_workbook_path, image_path, cell_address):
    # Open the existing Excel workbook that contains the VBA macro
    wb = xw.Book(input_workbook_path)
    
    # Access the first sheet (or specify the sheet name)
    sheet = wb.sheets[0]  # You can also use wb.sheets['SheetName'] if you want to specify the sheet by name
    
    # Specify the VBA macro name and parameters
    macro_name = "InsertPictureIntoCell"
    sheet.api.Run(macro_name, image_path, cell_address)

    # Save the workbook as a new file
    wb.save(output_workbook_path)
    wb.close()

def insert_picture_into_excel(input_workbook_path, output_workbook_path, image_path, cell_address):
    # Open the existing Excel workbook that contains the VBA macro
    wb = xw.Book(input_workbook_path)
    
    # Run the VBA macro
    macro = wb.macro('InsertImagesBasedOnSKU')  # Load the VBA macro
    macro()  # Call the macro with parameters
    
    # Save the workbook as a new file
    wb.save(output_workbook_path)
    wb.close()

def create_xlsm_with_vba_macro(vba_file_path, output_xlsm_path):
    # Create a new Excel workbook
    wb = xw.Book()
    
    # Save it as an .xlsm file to enable macros
    wb.save(output_xlsm_path)

    # Read the VBA macro from the file
    with open(vba_file_path, 'r') as vba_file:
        vba_code = vba_file.read()
    
    #printout the vba code
    print(vba_code)

    # Embed the VBA macro into the workbook
    wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)
    
    # Save the workbook with the embedded macro
    wb.save()
    wb.close()

if __name__ == "__main__":
    vba_file_path = 'vba/insert_images_macro.vba'  # Replace with the path to your .vba file
    output_xlsm_path = 'template_workbook.xlsm'    # Replace with the desired output file path

    create_xlsm_with_vba_macro(vba_file_path, output_xlsm_path)
    # input_workbook_path = 'demo.xlsm'  # Replace with the actual path to your .xlsm file
    # output_workbook_path = 'demo.xlsx'       # Replace with the desired output file path
    # image_path = 'test_images/1001.png'               # Replace with the actual path to your image
    # cell_address = 'B2'                                 # Replace with the desired cell address

    # # Insert picture into Excel and save as a new file
    # insert_picture_into_excel(input_workbook_path, output_workbook_path, image_path, cell_address)
