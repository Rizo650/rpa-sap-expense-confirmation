Sub Lookup()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim dict As Object
    Dim i As Long
    Dim key As String
    Dim targetSheetNames As Variant
    Dim sheetName As Variant

    ' Inisialisasi
    Set wsSource = ThisWorkbook.Sheets("ZFI_GRIR")
    Set dict = CreateObject("Scripting.Dictionary")

    ' Simpan data dari sheet ZFI_GRIR ke dictionary
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row
    For i = 2 To lastRowSource
        key = Trim(wsSource.Cells(i, "C").Value)
        If Len(key) > 0 And Not dict.exists(key) Then
            dict.Add key, wsSource.Cells(i, "Q").Value
        End If
    Next i

    ' Daftar sheet target: W1 dan W2
    targetSheetNames = Array("Detail W1", "Detail W2")

    ' Proses VLOOKUP untuk setiap sheet target
    For Each sheetName In targetSheetNames
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "AP").End(xlUp).Row

        For i = 2 To lastRowTarget
            key = Trim(wsTarget.Cells(i, "AP").Value)
            If dict.exists(key) Then
                wsTarget.Cells(i, "AK").Value = dict(key)
            End If
            ' Jika tidak ditemukan, nilai kolom AK dibiarkan
        Next i
    Next sheetName

End Sub
