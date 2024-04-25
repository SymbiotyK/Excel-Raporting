Private Sub Worksheet_Change(ByVal Target As Range)

    ActiveSheet.Unprotect ("")
    Dim rng As Range
    Dim rng2 As Range
    
    Set rng = Me.Range("C:C")
    Set rng2 = Me.Range("F:F")
    
    If Not Intersect(Target, rng) Is Nothing Then
    If Target.Value <> "" Then
        Target.Offset(0, 2).Value = Format(Now, "dd.mm.yyyy hh:mm:ss")
        Else
            Target.Offset(0, 2).ClearContents
        End If
    End If

    If Not Intersect(Target, rng2) Is Nothing Then
    If Target.Value <> "" Then
        Target.Offset(0, 1).Value = Format(Now, "dd.mm.yyyy hh:mm:ss")
        Else
            Target.Offset(0, 1).ClearContents
        End If
    End If
    ActiveSheet.Protect ("")
End Sub


