Sub ExportToPDF()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String

    ' Set the worksheet to be exported
    Set ws = ThisWorkbook.Sheets("druk")
    
    ' Define the file path and name
    filePath = "C:\Users\#UserID#\#OneDriveDirectory#\[...]"
    fileName = ws.Range("C3").Value & " - " & ws.Range("I3").Value & " - " & ws.Range("G3").Value
    
    ' Export the worksheet to PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath & fileName, Quality:=xlQualityStandard
End Sub




