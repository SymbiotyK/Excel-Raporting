
Sub CopyAndPasteData()
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim dest2Sheet As Worksheet
    Dim dateValue As Date
    Dim destRow As Long
    Dim destRow2 As Long
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim copyRange2 As Range
    Dim pasteRange2 As Range
    Dim proceed As VbMsgBoxResult
    
    Unlock_sheet
    
    proceed = MsgBox("UWAGA!!! Zapisuj dane TYLKO NA KONIEC zmiany. Czy chcesz zapisać dane?", vbQuestion + vbYesNo, "Potwierdzenie")
    
    If proceed = vbYes Then
        Set srcSheet = ThisWorkbook.Sheets("Karta")
        Set destSheet = ThisWorkbook.Sheets("Zapisane straty czasu")
        Set dest2Sheet = ThisWorkbook.Sheets("Zapisane sztuki")
        
        If (Application.WorksheetFunction.CountBlank(srcSheet.Range("J25:J53")) - 29) < 0 Then
           MsgBox "Proszę wprowadzić stracony czas dla każdej przyczyny przestoju.", vbExclamation, "UWAGA!"
           ThisWorkbook.Sheets("Karta").Activate
           Lock_sheet
           Exit Sub
         End If
         
        CopyAndPastePrint
        PrintSheet
        ExportToPDF
        
        dateValue = srcSheet.Range("G3").Value
        destRow = destSheet.Cells(destSheet.Rows.Count, 1).End(xlUp).Row + 1
        destRow2 = dest2Sheet.Cells(dest2Sheet.Rows.Count, 1).End(xlUp).Row + 1
        
        With destSheet.Range("A" & destRow)
            .Value = dateValue
            .Resize(1, 7).merge
            .HorizontalAlignment = xlCenter
            .NumberFormat = "dd/mm/yyyy"
        End With
        
        With dest2Sheet.Range("A" & destRow2)
            .Value = dateValue
            .Resize(1, 8).merge
            .HorizontalAlignment = xlCenter
            .NumberFormat = "dd/mm/yyyy"
        End With
        
        Set copyRange = srcSheet.Range("C25:I53")
        Set pasteRange = destSheet.Range("A" & destRow + 1)
    
        copyRange.Copy
        pasteRange.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    
        copyRange.ClearContents
        Application.CutCopyMode = False
        
        Set copyRange2 = srcSheet.Range("C12:J19")
        Set pasteRange2 = dest2Sheet.Range("A" & destRow2 + 1)
        
        copyRange2.Copy
        pasteRange2.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

        srcSheet.Range("D12:F19").ClearContents
        srcSheet.Range("H12:I19").ClearContents
        Application.CutCopyMode = False
    End If
    
    ThisWorkbook.Save
    ThisWorkbook.Sheets("Karta").Activate
    Lock_sheet
End Sub


