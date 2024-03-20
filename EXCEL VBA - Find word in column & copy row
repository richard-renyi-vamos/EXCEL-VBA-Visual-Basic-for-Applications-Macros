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
