
Function rowNumber(searchValue As Variant) As Integer
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Zapisane straty czasu")
    
    Dim searchRange As Range
    Dim firstDate As Range
    
    ' Looking for the row below the first date of the week
    Set searchRange = ws.Range("A:F")
    Set firstDate = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Returning the index of the row
    If firstDate Is Nothing Then
    Exit Function
    Else
    rowNumber = firstDate.Row
    End If
    
End Function

Sub DynamicSumIf()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Zapisane straty czasu")
    Dim result As Double
    Dim firstRow As Integer
    Dim lastRow As Integer
    
    Dim firstDate As Variant
    Dim lastDate As Variant
    
    ws.Unprotect Password:="god"
    
    firstDate = ws.Range("P3").Value
    lastDate = ws.Range("P4").Value
    
    firstRow = rowNumber(firstDate)
    lastRow = rowNumber(lastDate)
    
    If firstRow = 0 Then
        MsgBox "Nie można znaleźć pierwszej daty.", vbCritical
    Exit Sub
    End If
    
    If lastRow = 0 Then
        MsgBox "Nie można znaleźć ostatniej daty, dlatego patrze na ostatni pusty rząd", vbCritical
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    'Seting values for reasons from whole week
    For i = 3 To 10
        ws.Range("I" & i).Value = Application.WorksheetFunction.SumIf(ws.Range("B" & firstRow & ":B" & lastRow), ws.Range("H" & i).Value, ws.Range("F" & firstRow & ":F" & lastRow))
    Next i
    
    'Pareto sorting
    ws.ListObjects("Table1").Sort. _
        SortFields.Clear
    ws.ListObjects("Table1").Sort. _
        SortFields.Add2 Key:=Range("Table1[[#All],[czas]]"), SortOn:=xlSortOnValues _
        , Order:=xlDescending, DataOption:=xlSortNormal
    With ws.ListObjects("Table1"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("J3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=[@czas]"
    Range("J4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=[@czas]+R[-1]C"
    Range("J4").Select
    Selection.AutoFill Destination:=Range("J4:J10")
    Range("J4:J10").Select
    
    ThisWorkbook.Save
    ws.Protect Password:="god"
End Sub
