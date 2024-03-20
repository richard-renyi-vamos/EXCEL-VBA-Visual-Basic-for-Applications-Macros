

User
create an EXCEL VBA MACRO that finds every row in COLUMN A that contain the word "TEST" and copies those rows in a second excel table
ChatGPT
You can create a VBA macro in Excel to achieve this task. Below is a sample code that searches for rows containing the word "TEST" in Column A and copies them to a second Excel table:

vba
Copy code
Sub CopyRowsWithTest()

    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long, destLastRow As Long
    Dim i As Long
    
    ' Set the source and destination worksheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your source sheet
    Set wsDestination = ThisWorkbook.Sheets("Sheet2") ' Change "Sheet2" to the name of your destination sheet
    
    ' Find the last row in the source worksheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Find the last row in the destination worksheet
    destLastRow = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Loop through each row in Column A of the source worksheet
    For i = 1 To lastRow
        ' Check if the cell in Column A contains the word "TEST"
        If InStr(1, wsSource.Cells(i, 1).Value, "TEST", vbTextCompare) > 0 Then
            ' If found, copy the entire row to the destination worksheet
            wsSource.Rows(i).Copy Destination:=wsDestination.Rows(destLastRow)
            destLastRow = destLastRow + 1
        End If
    Next i
    
    MsgBox "Rows with 'TEST' copied successfully!", vbInformation
    
End Sub
To use this macro:

Press ALT + F11 to open the VBA editor.
Go to Insert > Module to insert a new module.
Copy and paste the above code into the module window.
Close the VBA editor.
You can now run the macro by going to Developer tab (if not visible, enable it via File > Options > Customize Ribbon > check Developer) and clicking on Macros, then selecting CopyRowsWithTest and clicking Run.
Make sure to adjust the sheet names ("Sheet1" and "Sheet2") to match your actual sheet names. This code assumes that the destination table starts from the first row of the destination sheet. Adjust the destination range if needed.





