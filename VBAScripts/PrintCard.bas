Sub PrintSheet()
    Dim ws As Worksheet
    Dim rng As Range
    Dim prn As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Druk")
    Set rng = ws.Range("C3:J53")
    prn = "PFF9"  ' Nazwa drukarki
    
    With ws.PageSetup
        .PrintArea = rng.Address
        .Orientation = xlLandscape
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PrintQuality = 600
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
    End With
    
    ws.PrintOut ActivePrinter:=prn
    On Error GoTo 0
    
    ThisWorkbook.Sheets("Karta").Activate
    
    MsgBox "Dokument został wydrukowany.", vbInformation, "Drukowanie"
    Exit Sub
    
ErrorHandler:
    MsgBox "Wystapil problem podczas drukowania, upewnij sie czy drukarka jest sprawna.", vbExclamation, "Error"
    On Error GoTo 0
End Sub
