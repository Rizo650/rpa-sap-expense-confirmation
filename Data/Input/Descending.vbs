Sub Descending()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim targetSheets As Variant
    Dim sheetName As Variant

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    targetSheets = Array("Detail W1", "Detail W2")

    For Each sheetName In targetSheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        If ws Is Nothing Then
            On Error GoTo 0
            GoTo NextSheet
        End If
        On Error GoTo 0

        ' Tentukan baris awal berdasarkan nama sheet
        Select Case sheetName
            Case "Detail W1", "Detail W2"
                startRow = 34
            Case "Import Expense"
                startRow = 5
            Case Else
                startRow = 2 ' default
        End Select

        ' Cari baris terakhir di kolom J
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        If lastRow < startRow Then GoTo NextSheet ' Skip jika tidak ada data untuk disortir

        ' Lakukan sorting descending berdasarkan kolom J
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("J" & startRow & ":J" & lastRow), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange ws.Range("A" & startRow & ":AW" & lastRow)
            .Header = xlNo
            .Apply
        End With

NextSheet:
        Set ws = Nothing
    Next sheetName

    ThisWorkbook.Save
    DoEvents
End Sub
