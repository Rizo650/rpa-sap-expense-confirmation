Sub Highlight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim descTextK As String
    Dim descTextAK As String
    Dim val As Double
    Dim targetSheets As Variant
    Dim sheetName As Variant

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    targetSheets = Array("Detail W1", "Detail W2", "Import Expense")

    For Each sheetName In targetSheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        If ws Is Nothing Then
            On Error GoTo 0
            GoTo NextSheet
        End If
        On Error GoTo 0

        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

        For i = 2 To lastRow
            descTextK = LCase(ws.Cells(i, "K").Value)
            descTextAK = LCase(ws.Cells(i, "AK").Value)
            val = 0
            If IsNumeric(ws.Cells(i, "AR").Value) Then val = ws.Cells(i, "AR").Value

            If InStr(descTextK, "roll") > 0 _
                Or InStr(descTextK, "sleeve ring") > 0 _
                Or InStr(descTextK, "arbor") > 0 _
                Or InStr(descTextAK, "roll") > 0 _
                Or InStr(descTextAK, "sleeve ring") > 0 _
                Or InStr(descTextAK, "arbor") > 0 _
                Or val > 100000000 Then

                ws.Range(ws.Cells(i, "B"), ws.Cells(i, "AW")).Interior.Color = RGB(255, 255, 0)
            End If
        Next i

NextSheet:
        Set ws = Nothing
    Next sheetName

    ThisWorkbook.Save
    DoEvents
End Sub
