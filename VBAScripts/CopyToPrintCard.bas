Sub CopyAndPastePrint()

    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim srcRange As Range
    Dim destRange As Range
      
    Set srcSheet = ThisWorkbook.Sheets("Karta")
    Set srcRange = srcSheet.Range("A1:L55")
    
    Set destSheet = ThisWorkbook.Sheets("Druk")
    Set destRange = destSheet.Range("A1:L55")
    
    srcRange.Copy
    destRange.PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False

End Sub
