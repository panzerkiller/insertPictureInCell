import os
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

def insert_picture_into_excel(input_workbook_path, output_workbook_path, image_path):
    # Open the existing Excel workbook that contains the VBA macro
    wb = xw.Book(input_workbook_path)
    
    # Run the VBA macro
    macro = wb.macro('Sheet1.InsertImagesBasedOnSKU')  # Load the VBA macro
    print(image_path)
    #print the macro
    print(macro)
    macro()  # Call the macro with parameters
    
    # Save the workbook as a new file
    wb.save(output_workbook_path)
    wb.close()

def check_macro_loaded(wb, macro_name):
    try:
        # Attempt to load and run the macro without parameters
        macro = wb.macro(macro_name)
        
        # Run the macro (if it requires parameters, you can pass dummy data here)
        macro()
        
        print(f"The macro '{macro_name}' is correctly loaded and executed.")
        return True
    except Exception as e:
        print(f"Failed to load or run the macro '{macro_name}'. Error: {e}")
        return False

def create_xlsm_with_vba_macro(vba_file_path, image_folder, output_xlsm_path):
    # Start an instance of Excel and hide the window
    app = xw.App(visible=False)
    # Create a new Excel workbook
    wb = xw.Book()
    
    #rename the sheet name to 'Images'
    wb.sheets[0].name = 'Images'

    # Save it as an .xlsm file to enable macros
    wb.save(output_xlsm_path)

    # Read the VBA macro from the file
    with open(vba_file_path, 'r') as vba_file:
        vba_code = vba_file.read()
    
    # Replace the <image_folder> placeholder with the actual image folder path
    vba_code = vba_code.replace('<image_folder>', image_folder)
    #printout the vba code
    #print(vba_code)

    # Embed the VBA macro into the workbook
    wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)
    
    # Save the workbook with the embedded macro
    wb.save()
    wb.close()
    # Quit the Excel application
    app.quit()

def add_sku_to_xlsm(xlsm_file_path, png_folder_path, sheet_name='Sheet1'):
    # Start an instance of Excel and hide the window
    app = xw.App(visible=False)

    # Open the existing Excel workbook
    wb = xw.Book(xlsm_file_path)
    
    # Access the specified sheet (or use the default sheet name)
    sheet = wb.sheets[sheet_name]
    
    # Start adding file names from the second row (first row for headers)
    start_row = 2
    
    # Set the header for the first column
    sheet.range('A1').value = 'SKU'
    
    # Iterate through the PNG files in the specified folder
    for i, file_name in enumerate(os.listdir(png_folder_path)):
        if file_name.endswith('.png'):
            # Remove the file extension to get the SKU
            sku = os.path.splitext(file_name)[0]
            
            # Write the SKU into the corresponding cell in the first column
            sheet.range(f'A{start_row + i}').value = sku
    
    # Save the workbook and close it
    wb.save()
    wb.close()

    # Quit the Excel application
    app.quit()

def insert_folder_path(folder_path, cell_address, workbook_path):
    # Open the workbook
    wb = xw.Book(workbook_path)
    
    # Load the macro
    macro = wb.macro('InsertFolderPathIntoCell')
    
    # Run the macro with the folder path and cell address
    macro(folder_path, cell_address)
    
    # Save and close the workbook
    wb.save()
    wb.close()

def fill_cell_with_hello_world(workbook_path):
    # Open the workbook
    wb = xw.Book(workbook_path)
    
    # Load the macro
    macro = wb.macro('Sheet1.FillCellWithHelloWorld')
    
    # Run the macro
    print(f"Running macro: {macro}")
    macro()  # Execute the macro
    
    # Save the workbook
    print("Saving the workbook...")
    wb.save()
    
    # Close the workbook
    wb.close()

if __name__ == "__main__":
    # #vba_file_path = 'vba/insert_images_macro.vba'  # Replace with the path to your .vba file
    workbook_path = 'template_workbook.xlsm'    # Replace with the desired output file path
    
    # #create_xlsm_with_vba_macro(vba_file_path, 'test_images', output_xlsm_path)

    add_sku_to_xlsm(workbook_path, 'test_images', 'Images')    
    # # input_workbook_path = 'demo.xlsm'  # Replace with the actual path to your .xlsm file
    output_workbook_path = 'demo.xlsx'       # Replace with the desired output file path
    image_path = '/Users/haoli/repo/insertPictureInCell/test_images'               # Replace with the actual path to your image
    # # cell_address = 'B2'                                 # Replace with the desired cell address

    # # # Insert picture into Excel and save as a new file
    insert_picture_into_excel(workbook_path, output_workbook_path, image_path)


    # workbook_path = 'template_workbook.xlsm'  # Replace with your actual .xlsm file path
    #macro_name = 'InsertImagesBasedOnSKU'  # Replace with your actual macro name
    
    # Open the workbook
    wb = xw.Book(workbook_path)
    
    # Check if the macro is loaded correctly
    #macro_loaded = check_macro_loaded(wb, macro_name)
    
   

    folder_path = 'test_images'  # Replace with your actual folder path
    cell_address = 'A1'  # Replace with the desired cell address

    #insert_folder_path(folder_path, cell_address, workbook_path)

    #fill_cell_with_hello_world(workbook_path)
    