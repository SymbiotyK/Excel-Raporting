Sub HighlightRows()
    Dim startRow As Long
    Dim endRow As Long
    Dim ws As Worksheet
    Dim rng As Range
    Dim highlightRange As Range

    ' Set the worksheet
    Set ws = ActiveSheet

    Unlock_sheet

    If ws.Name = "Zapisane straty czasu" Then
        ws.Range("A2:F" & ws.Cells(ws.Rows.Count, 6).End(xlUp).Row).Interior.Color = RGB(248, 203, 173) ' Pink color
    Else
        ws.Range("A2:G" & ws.Cells(ws.Rows.Count, 7).End(xlUp).Row).Interior.Color = RGB(248, 203, 173) ' Pink color
    End If
    
        'check if there is something in R3 R4 ; L3 L4
    If IsError(ws.Range("R3").Value) Or IsError(ws.Range("L3").Value) Then
    MsgBox "Nie można znaleźć pierwszej daty", vbCritical
    Exit Sub
    End If
    
    If IsError(ws.Range("R4").Value) Or IsError(ws.Range("L4").Value) Then
        MsgBox "Nie można znaleźć ostatniej daty, dlatego patrze na ostatni wiersz", vbCritical
        If ws.Name = "Zapisane straty czasu" Then
            startRow = ws.Range("R3").Value
            endRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Else
            startRow = ws.Range("L3").Value
            endRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        End If
    Else
    ' Read start and end row numbers from cells R3 R4 ; L3 L4
        If ws.Name = "Zapisane straty czasu" Then
            startRow = ws.Range("R3").Value
            endRow = ws.Range("R4").Value
        Else
            startRow = ws.Range("L3").Value
            endRow = ws.Range("L4").Value
        End If
    End If
    ' Set the range to be highlighted
    If ws.Name = "Zapisane straty czasu" Then
        Set rng = ws.Range("A" & startRow & ":F" & endRow)
    Else
        Set rng = ws.Range("A" & startRow & ":G" & endRow)
    End If

    ' Highlight the range with a desired color
    rng.Interior.Color = RGB(255, 255, 0) ' Yellow color

    ' Scroll to the first highlighted row for better visibility
    If Not rng Is Nothing Then
        ws.Activate
        rng.Cells(1).Activate
    End If
    
    Lock_sheet
End Sub


