Sub PairTablesBasedOnEmailAndCopyToNewWorkbook()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim tbl1 As ListObject, tbl2 As ListObject
    Dim rng1 As Range, rng2 As Range
    Dim emailColumn1 As Range, emailColumn2 As Range
    Dim emailDict As Object
    Dim email As String, row1 As Long, row2 As Long
    Dim newWorkbook As Workbook
    Dim newRow As Long
    
    ' Open a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Set references to the worksheets
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your first worksheet
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' Change "Sheet2" to the name of your second worksheet
    Set ws3 = newWorkbook.Sheets(1) ' Assumes the new workbook only has one sheet
    
    ' Assuming your tables have headers, change the range accordingly if they don't
    Set tbl1 = ws1.ListObjects("Table1") ' Change "Table1" to the name of your first table
    Set tbl2 = ws2.ListObjects("Table2") ' Change "Table2" to the name of your second table
    
    ' Assuming the email addresses are in the first column of each table
    Set emailColumn1 = tbl1.ListColumns(1).DataBodyRange
    Set emailColumn2 = tbl2.ListColumns(1).DataBodyRange
    
    ' Create a dictionary to store pairs of email addresses and corresponding row numbers
    Set emailDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the first table and populate the dictionary with email addresses and row numbers
    For row1 = 1 To emailColumn1.Rows.Count
        email = emailColumn1.Cells(row1, 1).Value
        If Not emailDict.exists(email) Then
            emailDict(email) = row1
        End If
    Next row1
    
    ' Loop through the second table and pair rows based on matching email addresses
    newRow = 1
    For row2 = 1 To emailColumn2.Rows.Count
        email = emailColumn2.Cells(row2, 1).Value
        If emailDict.exists(email) Then
            ' If a matching email address is found in the first table, copy the paired rows to the new workbook
            tbl1.DataBodyRange.Rows(emailDict(email)).Copy ws3.Cells(newRow, 1)
            tbl2.DataBodyRange.Rows(row2).Copy ws3.Cells(newRow, tbl1.ListColumns.Count + 1)
            newRow = newRow + 1
        End If
    Next row2
    
    ' Clean up
    Set ws1 = Nothing
    Set ws2 = Nothing
    Set ws3 = Nothing
    Set tbl1 = Nothing
    Set tbl2 = Nothing
    Set emailColumn1 = Nothing
    Set emailColumn2 = Nothing
    Set emailDict = Nothing
    Set newWorkbook = Nothing
    
    MsgBox "Pairing and copying completed!", vbInformation
End Sub
