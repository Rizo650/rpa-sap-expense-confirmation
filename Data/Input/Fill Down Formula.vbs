Sub FillDownFormulaDynamic()
    Dim ws As Worksheet
    Dim tempLastRow As Long
    Dim lastRow As Long
    Dim startCell As Range

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Detail W1", "Detail W2"
                Set startCell = ws.Range("AR34")
                tempLastRow = ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row
            Case "Import Expense"
                Set startCell = ws.Range("AR5")
                tempLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
            Case Else
                ' Sheet lain tidak diproses
                GoTo NextSheet
        End Select

        lastRow = Application.WorksheetFunction.Max(tempLastRow, startCell.Row)

        If lastRow > startCell.Row Then
            ws.Range(startCell, ws.Cells(lastRow, startCell.Column)).FillDown
        End If

NextSheet:
    Next ws
End Sub

