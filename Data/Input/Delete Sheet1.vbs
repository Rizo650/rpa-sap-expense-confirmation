Sub DeleteSheet1()
    Application.DisplayAlerts = False ' Supaya tidak muncul konfirmasi
    On Error Resume Next ' Supaya tidak error kalau Sheet1 tidak ada
    ThisWorkbook.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
End Sub