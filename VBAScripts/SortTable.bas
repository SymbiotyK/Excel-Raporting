Sub tableSorting()
    ThisWorkbook.Sheets("Zapisane sztuki").Unprotect Password:="#password"
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=2, Criteria1:="<>0", Operator:=xlFilterValues
    ActiveWorkbook.Worksheets("Zapisane sztuki").ListObjects("Table2").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Zapisane sztuki").ListObjects("Table2").Sort. _
        SortFields.Add2 Key:=Range("Table2[[#All],[Sztuki OK]]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Zapisane sztuki").ListObjects("Table2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ThisWorkbook.Sheets("Zapisane sztuki").Protect Password:="#password"
End Sub
Sub scrollTable()
    ThisWorkbook.Sheets("Zapisane sztuki").Unprotect Password:="#password"
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=2
    ThisWorkbook.Sheets("Zapisane sztuki").Protect Password:="#password"
End Sub
