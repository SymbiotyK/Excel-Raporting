Private Sub CommandButton1_Click()
    
    Application.ScreenUpdating = False
    Dim wsAnaliza As Worksheet
    Dim wsDane As Worksheet
    Dim fDate As Variant
    Dim lDate As Variant
    Dim whereFirstSort As Range
    Dim whereLastSort As Range
    Dim whereFirstLoad As Range
    Dim whereLastLoad As Range
    Dim whereFirstSearch As Range
    Dim filePath As String
    Dim myTable As ListObject
    Dim whereScrDataColumn As Integer
    
    Set wsAnaliza = ThisWorkbook.Sheets("Analiza")
    Set wsDane = ThisWorkbook.Sheets("Dane")
    
    fDate = wsAnaliza.Range("B3").Value
    lDate = wsAnaliza.Range("B4").Value
    
    ' #1 Set range references
    Set whereFirstSort = wsDane.Range("C3")
    Set whereLastSort = wsDane.Range("C10")
    Set whereFirstLoad = wsDane.Range("B3")
    Set whereLastLoad = wsDane.Range("B10")
    Set whereFirstSearch = wsDane.Range("A3")
    Set myTable = wsDane.ListObjects("Table1")
    whereScrDataColumn = 7
    filePath = "jedox_AWS1"
    fillTable whereFirstSearch, whereFirstLoad, whereLastLoad, whereFirstSort, whereLastSort, fDate, lDate, filePath, myTable, whereScrDataColumn
    
    ' #2 Set range references
    Set whereFirstSort = wsDane.Range("C14")
    Set whereLastSort = wsDane.Range("C21")
    Set whereFirstLoad = wsDane.Range("B14")
    Set whereLastLoad = wsDane.Range("B21")
    Set whereFirstSearch = wsDane.Range("A14")
    Set myTable = wsDane.ListObjects("Table2")
    whereScrDataColumn = 7
    filePath = "jedox_AWS2"
    fillTable whereFirstSearch, whereFirstLoad, whereLastLoad, whereFirstSort, whereLastSort, fDate, lDate, filePath, myTable, whereScrDataColumn
    
    ' #3 Set range references
    Set whereFirstSort = wsDane.Range("C25")
    Set whereLastSort = wsDane.Range("C32")
    Set whereFirstLoad = wsDane.Range("B25")
    Set whereLastLoad = wsDane.Range("B32")
    Set whereFirstSearch = wsDane.Range("A25")
    Set myTable = wsDane.ListObjects("Table3")
    whereScrDataColumn = 7
    filePath = "jedox_LAGMontaz"
    fillTable whereFirstSearch, whereFirstLoad, whereLastLoad, whereFirstSort, whereLastSort, fDate, lDate, filePath, myTable, whereScrDataColumn
    
    ' #4 Set range references
    Set whereFirstSort = wsDane.Range("J3")
    Set whereLastSort = wsDane.Range("J10")
    Set whereFirstLoad = wsDane.Range("I3")
    Set whereLastLoad = wsDane.Range("I10")
    Set whereFirstSearch = wsDane.Range("H3")
    Set myTable = wsDane.ListObjects("Table4")
    whereScrDataColumn = 7
    filePath = "jedox_LPSStacja"
    fillTable whereFirstSearch, whereFirstLoad, whereLastLoad, whereFirstSort, whereLastSort, fDate, lDate, filePath, myTable, whereScrDataColumn
    
    ' #5 Set range references
    Set whereFirstSort = wsDane.Range("J14")
    Set whereLastSort = wsDane.Range("J21")
    Set whereFirstLoad = wsDane.Range("I14")
    Set whereLastLoad = wsDane.Range("I21")
    Set whereFirstSearch = wsDane.Range("H14")
    Set myTable = wsDane.ListObjects("Table5")
    whereScrDataColumn = 7
    filePath = "jedox_LPSMontaz"
    fillTable whereFirstSearch, whereFirstLoad, whereLastLoad, whereFirstSort, whereLastSort, fDate, lDate, filePath, myTable, whereScrDataColumn
    
    
    wsAnaliza.Activate
    Application.ScreenUpdating = True
    MsgBox "Kopiowanie czasu zakończone", vbInformation
    
End Sub

Private Sub CommandButton2_Click()

    Application.ScreenUpdating = False
    Dim whereFirstLoad As Range
    Dim fDate As Variant
    Dim lDate As Variant
    Dim filePath As String
    Dim oknTab As Integer

    fDate = ThisWorkbook.Sheets("Analiza").Range("B3").Value
    lDate = ThisWorkbook.Sheets("Analiza").Range("B4").Value
    
    oknTab = 6
    
    '#1 OK
    filePath = "jedox_AWS1"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("I25")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#2 OK
    filePath = "jedox_AWS2"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("I26")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#3 OK
    filePath = "jedox_LAGMontaz"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("I27")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#4 OK
    filePath = "jedox_LPSStacja"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("I28")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#5 OK
    filePath = "jedox_LPSMontaz"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("I29")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    oknTab = 7

    '#1 NOK
    filePath = "jedox_AWS1"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("J25")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#2 NOK
    filePath = "jedox_AWS2"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("J26")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#3 NOK
    filePath = "jedox_LAGMontaz"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("J27")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#4 NOK
    filePath = "jedox_LPSStacja"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("J28")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab
    
    '#5 NOK
    filePath = "jedox_LPSMontaz"
    Set whereFirstLoad = ThisWorkbook.Sheets("Dane").Range("J29")
    sumAmount whereFirstLoad, fDate, lDate, filePath, oknTab


    ThisWorkbook.Sheets("Analiza").Activate
    Application.ScreenUpdating = True
    MsgBox "Kopiowanie sztuk zakończone", vbInformation

End Sub

Private Sub CommandButton3_Click()

    Application.ScreenUpdating = False
    ActiveSheet.ChartObjects("LAGMon").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.ChartObjects("LAGSta1").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.ChartObjects("LAGSta2").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.ChartObjects("LPSMon").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.ChartObjects("LPSSta").Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("AO33:AZ59").Select
    Selection.PrintOut Copies:=1, Collate:=True
    
    ThisWorkbook.Sheets("Analiza").Activate
    Application.ScreenUpdating = True
    MsgBox "Drukowanie zakończone", vbInformation

End Sub
