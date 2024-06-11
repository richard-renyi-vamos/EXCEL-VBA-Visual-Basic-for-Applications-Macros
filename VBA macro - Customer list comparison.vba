Sub CheckCustomersInSharePoint()
    Dim wsOriginal As Worksheet
    Dim wsSharePoint As Worksheet
    Dim originalLastRow As Long
    Dim sharePointLastRow As Long
    Dim originalCustomer As String
    Dim foundCustomer As Range
    Dim i As Long

    ' Define the SharePoint file path and worksheet name
    Dim sharePointFilePath As String
    Dim sharePointSheetName As String
    sharePointFilePath = "https://yoursharepointsite.com/yourfolder/yourfile.xlsx"
    sharePointSheetName = "Sheet1" ' Change this to the actual sheet name in the SharePoint file

    ' Set the original worksheet
    Set wsOriginal = ThisWorkbook.Sheets("Sheet1") ' Change this to the actual sheet name in the original file

    ' Add a new column to the original worksheet
    wsOriginal.Columns("B").Insert Shift:=xlToRight
    wsOriginal.Cells(1, 2).Value = "Customer Check"

    ' Get the last row with data in the original worksheet
    originalLastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row

    ' Open the SharePoint workbook
    Dim sharePointWorkbook As Workbook
    Set sharePointWorkbook = Workbooks.Open(sharePointFilePath, ReadOnly:=True)

    ' Set the SharePoint worksheet
    Set wsSharePoint = sharePointWorkbook.Sheets(sharePointSheetName)

    ' Get the last row with data in the SharePoint worksheet
    sharePointLastRow = wsSharePoint.Cells(wsSharePoint.Rows.Count, "A").End(xlUp).Row

    ' Loop through each customer in the original worksheet
    For i = 2 To originalLastRow ' Assuming the first row is headers
        originalCustomer = wsOriginal.Cells(i, 1).Value
        If originalCustomer <> "" Then
            ' Check if the customer is in the SharePoint worksheet
            Set foundCustomer = wsSharePoint.Columns("A").Find(What:=originalCustomer, LookAt:=xlWhole)
            If Not foundCustomer Is Nothing Then
                ' If found, add the customer name to the new column
                wsOriginal.Cells(i, 2).Value = originalCustomer
            Else
                ' If not found, leave the cell empty
                wsOriginal.Cells(i, 2).Value = ""
            End If
        End If
    Next i

    ' Close the SharePoint workbook without saving
    sharePointWorkbook.Close SaveChanges:=False

    MsgBox "Customer check completed!"

End Sub
