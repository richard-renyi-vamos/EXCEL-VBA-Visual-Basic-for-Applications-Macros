Sub MergeExcelFiles()
    Dim SharepointURL1 As String
    Dim SharepointURL2 As String
    Dim LocalPath1 As String
    Dim LocalPath2 As String
    Dim DestinationPath As String
    Dim WB1 As Workbook
    Dim WB2 As Workbook
    Dim DestWB As Workbook
    Dim WS As Worksheet
    Dim CopyRange As Range
    Dim LastRow As Long

    ' Define SharePoint URLs
    SharepointURL1 = "https://yoursharepointsite/documents/ExcelFile1.xlsx"
    SharepointURL2 = "https://yoursharepointsite/documents/ExcelFile2.xlsx"
    
    ' Define local paths for temporary download
    LocalPath1 = Environ("TEMP") & "\ExcelFile1.xlsx"
    LocalPath2 = Environ("TEMP") & "\ExcelFile2.xlsx"
    
    ' Define the destination path
    DestinationPath = Environ("TEMP") & "\MergedExcelFile.xlsx"

    ' Download the files from SharePoint
    DownloadFileFromSharePoint SharepointURL1, LocalPath1
    DownloadFileFromSharePoint SharepointURL2, LocalPath2

    ' Open the downloaded files
    Set WB1 = Workbooks.Open(LocalPath1)
    Set WB2 = Workbooks.Open(LocalPath2)

    ' Create a new workbook for the merged data
    Set DestWB = Workbooks.Add

    ' Copy data from the first workbook
    For Each WS In WB1.Sheets
        WS.Copy After:=DestWB.Sheets(DestWB.Sheets.Count)
    Next WS

    ' Copy data from the second workbook
    For Each WS In WB2.Sheets
        WS.Copy After:=DestWB.Sheets(DestWB.Sheets.Count)
    Next WS

    ' Save the merged workbook
    DestWB.SaveAs DestinationPath

    ' Close the workbooks
    WB1.Close False
    WB2.Close False
    DestWB.Close False

    ' Inform the user
    MsgBox "Files merged successfully and saved at " & DestinationPath, vbInformation
End Sub

Sub DownloadFileFromSharePoint(SharepointURL As String, LocalPath As String)
    Dim xmlHTTP As Object
    Dim bStrm As Object
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHTTP.Open "GET", SharepointURL, False
    xmlHTTP.Send
    If xmlHTTP.Status = 200 Then
        Set bStrm = CreateObject("ADODB.Stream")
        bStrm.Type = 1 'binary
        bStrm.Open
        bStrm.Write xmlHTTP.responseBody
        bStrm.SaveToFile LocalPath, 2 'overwrite
        bStrm.Close
    Else
        MsgBox "Failed to download file from " & SharepointURL, vbCritical
    End If
End Sub
