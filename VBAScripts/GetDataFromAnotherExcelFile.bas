Function sumAmount(whereFirstLoad As Range, fDate As Variant, lDate As Variant, filePath As String, oknTab As Integer)
    
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim workBookPath As Workbook
    Dim firstRow As Integer
    Dim lastRow As Integer
    Dim i As Integer
    Dim LoginName As String
    LoginName = Environ("Username")
    
    Set wsDest = ThisWorkbook.Sheets("Dane")
    
    On Error Resume Next
    Set workBookPath = Workbooks.Open("C:\Users\#UserID#\#OneDriveDirectory#\[...]\" & filePath, ReadOnly:=True)
    On Error GoTo 0
    
    If workBookPath Is Nothing Then
        MsgBox "Plik nie jest otwarty.", vbCritical
        Exit Function
    End If
    
    Set wsSource = workBookPath.Sheets("Zapisane sztuki")
    
    firstRow = RowNumber(fDate, filePath, "Zapisane sztuki", "A:H") + 1
    lastRow = RowNumber(lDate, filePath, "Zapisane sztuki", "A:H") - 1
    
    If firstRow <= 0 Then
        MsgBox filePath & " - Nie można znaleźć pierwszej daty.", vbCritical
        Exit Function
    End If
    
    If lastRow <= 0 Then
        'MsgBox filePath & " - Nie można znaleźć ostatniej daty, dlatego patrze na ostatni pusty rząd", vbCritical
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    End If
    
    wsDest.Cells(whereFirstLoad.Row, whereFirstLoad.Column).Value = ""
    
    For i = firstRow To lastRow
        wsDest.Cells(whereFirstLoad.Row, whereFirstLoad.Column).Value = wsDest.Cells(whereFirstLoad.Row, whereFirstLoad.Column).Value + wsSource.Cells(i, oknTab).Value
    Next i
    
    workBookPath.Close SaveChanges:=False
    
End Function


Function loadData(filePath As String, whereFirstLoad As Range, whereLastLoad As Range, whereFirstSearch As Range, firstRow As Integer, lastRow As Integer, whereData As Integer)

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim i As Integer
    Dim destColumn As Integer
    Dim destSearchColumn As Integer
    
    ' Check if the workbook is open
    On Error Resume Next
    Set wsSource = Workbooks(filePath).Sheets("Zapisane straty czasu")
    Set wsDest = ThisWorkbook.Sheets("Dane")
    On Error GoTo 0
    
    If wsSource Is Nothing Or wsDest Is Nothing Then
        MsgBox "Nie znaleziono pliku.", vbCritical
        Exit Function
    End If
    
    destColumn = whereFirstLoad.Column
    destSearchColumn = whereFirstSearch.Column
    
    For i = whereFirstLoad.Row To whereLastLoad.Row
        Dim criteria As Variant
        criteria = wsDest.Cells(i, destSearchColumn).Value
        wsDest.Cells(i, destColumn).Value = SumIf(wsSource, firstRow, lastRow, criteria, whereData)
    Next i

End Function

Function SumIf(wsSource As Worksheet, firstRow As Integer, lastRow As Integer, criteria As Variant, whereData As Integer) As Double
    Dim sumResult As Double
    sumResult = 0
    
    For i = firstRow To lastRow
        If wsSource.Cells(i, 2).Value = criteria Then
            sumResult = sumResult + wsSource.Cells(i, whereData).Value
        End If
    Next i
    
    SumIf = sumResult
End Function

Function sortPareto(myTable As ListObject, firstRow As Range, lastRow As Range)

    myTable.Sort.SortFields.Clear
    myTable.Sort.SortFields.Add2 Key:=myTable.ListColumns("czas").Range, SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal
    
    With myTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    firstRow.FormulaR1C1 = "=[@czas]"
    firstRow.Offset(1, 0).FormulaR1C1 = "=[@czas]+R[-1]C"
    firstRow.Offset(1, 0).AutoFill Destination:=Range(firstRow.Offset(1, 0), lastRow.Offset(0, 0))

End Function

Function RowNumber(searchValue As Variant, filePath As String, sheetName As String, sourceRange As String) As Long

    Dim ws As Worksheet
    Dim path As String
    
    ' Check if the workbook is open
    On Error Resume Next
    Set ws = Workbooks(filePath).Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Plik nie jest otwarty.", vbCritical
        RowNumber = -1 ' Return -1 to indicate an error
        Exit Function
    End If
    
    ' Looking for the row below the first date of the week
    Set searchRange = ws.Range(sourceRange)
    Set firstDate = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if firstDate is Nothing (no match found)
    If firstDate Is Nothing Then
        RowNumber = -1 ' Value not found
    Else
        RowNumber = firstDate.Row
    End If
    
End Function
Sub fillTable(whereFirstSearch As Range, whereFirstLoad As Range, whereLastLoad As Range, whereFirstSort As Range, whereLastSort As Range, fDate As Variant, lDate As Variant, filePath As String, myTable As ListObject, whereData As Integer)

    Dim ws As Worksheet
    Dim path As String
    Dim workBookPath As Workbook
    Dim firstRow As Integer
    Dim lastRow As Integer
    Dim LoginName As String
    LoginName = Environ("Username")
    
    path = "C:\Users\#UserID#\#OneDriveDirectory#\[...]\" & filePath
    Set workBookPath = Workbooks.Open(path, ReadOnly:=True)
    
    ' Check if the workbook is open
    On Error Resume Next
    Set ws = Workbooks(filePath).Sheets("Zapisane straty czasu")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Plik nie jest otwarty.", vbCritical
        Exit Sub
    End If
    
    firstRow = RowNumber(fDate, filePath, "Zapisane straty czasu", "A:G") + 1
    lastRow = RowNumber(lDate, filePath, "Zapisane straty czasu", "A:G") - 1
    
    If firstRow <= 0 Then
        MsgBox filePath & " - Nie można znaleźć pierwszej daty.", vbCritical
        Exit Sub
    End If
    
    If lastRow <= 0 Then
        'MsgBox filePath & " - Nie można znaleźć ostatniej daty, dlatego patrze na ostatni pusty rząd", vbCritical
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    loadData filePath, whereFirstLoad, whereLastLoad, whereFirstSearch, firstRow, lastRow, whereData
    sortPareto myTable, whereFirstSort, whereLastSort
    
    'MsgBox "First Row: " & firstRow & vbCrLf & "Last Row: " & lastRow
    workBookPath.Close SaveChanges:=False

End Sub
