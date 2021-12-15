Sub macro()

    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("template")
    Dim kiten As Date
    kiten = ws1.Range("N5").Value
    Dim k As Long, hiduke As Date
    For k = 0 To 6
        hiduke = DateAdd("d", k, kiten)
        ws1.Range("D5").Offset(0, k).Value = Format(hiduke, "yyyy/mm/dd (aaa)")
        
        If Weekday(hiduke) = 1 Then
            ws1.Range("D5:D13").Offset(0, k).Interior.ColorIndex = 38
        ElseIf Weekday(hiduke) = 7 Then
            ws1.Range("D5:D13").Offset(0, k).Interior.ColorIndex = 34
        End If
    Next
    
    Dim filename As String
    filename = Format(kiten, "yyyy_mm_dd") & "_" & Format(hiduke, "yyyy_mm_dd") & "チェックリスト"
    ws1.Name = filename
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\" & filename
    
End Sub