Sub ConvertToDatesInRange()
    Dim cell As Range
    Dim rng As Range
    
    ' Define the range A1 to A5
    Set rng = Range("A1:A5")
    
    ' Loop through each cell in the defined range
    For Each cell In rng
        ' Check if the cell contains a numeric value
        If IsNumeric(cell.Value) Then
            ' Convert the numeric value to a date
            cell.Value = DateAdd("d", cell.Value - 1, "1899-12-31")
            ' Change the cell format to display as a date
            cell.NumberFormat = "mm/dd/yyyy"
        Else
            ' Notify if the selected cell does not contain a numeric value
            MsgBox "Cell " & cell.Address & " does not contain a numeric value.", vbExclamation
        End If
    Next cell
End Sub
