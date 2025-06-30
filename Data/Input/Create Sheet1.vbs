Sub BuatSheet1()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sheet1")
    On Error GoTo 0
    
    ' Jika Sheet1 belum ada, buat baru
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Sheet1"
    End If
End Sub