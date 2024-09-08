' insert_images_macro.vba

Sub InsertImagesBasedOnSKU()
    Dim ws As Worksheet
    Dim cell As Range
    Dim imgPath As String
    Dim imgName As String
    Dim sku As String
    Dim imgFolder As String

    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Images") ' Change to your sheet name

    ' Define the folder where the images are stored
    imgFolder = "<image_folder>" ' Update this to your image folder

    ' Loop through each cell in the first column (assume SKU is in column A)
    ' For Each cell In ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    For Each cell In ws.Range("A2:A4") ' & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)    
        sku = cell.Value
        imgName = sku & ".png"
        imgPath = imgFolder & "\" & imgName
        
        ' Check if the image file exists
        If Dir(imgPath) <> "" Then
            ' Select the cell where the image should be inserted
            cell.Offset(0, 1).Select
            
            ' Insert the picture in the selected cell
            Selection.InsertPictureInCell (imgPath)
            
            ' Write the image name to the third column
            cell.Offset(0, 2).Value = imgName
        Else
            ' If the image file doesn't exist, notify the user
            cell.Offset(0, 1).Value = "Image not found"
            cell.Offset(0, 2).Value = "N/A"
        End If
    Next cell
End Sub
