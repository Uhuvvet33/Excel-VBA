Option Explicit
Sub VerileriBirlestir()
    Dim ws As Worksheet
    Dim hedefSayfa As Worksheet
    Dim sonSatir As Long
    Dim i As Long
    
    ' Hedef sayfanın adını ve hedef satır numarasını ayarlayın
    Set hedefSayfa = ThisWorkbook.Sheets("topla")
    sonSatir = hedefSayfa.Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    ' Tüm sayfaları dolaşarak verileri birleştirin
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "topla" Then
            ' İlgili veri aralığını ayarlayın (örneğin B7:D7)
            For i = 4 To 6 'verinin alınacağı ilk satır
                hedefSayfa.Cells(sonSatir, 1).Value = ws.Cells(i, 1).Value
                hedefSayfa.Cells(sonSatir, 2).Value = ws.Cells(i, 2).Value
                hedefSayfa.Cells(sonSatir, 3).Value = ws.Cells(i, 3).Value
                'hedefSayfa.Cells(sonSatir, 4).Value = ws.Cells(i, 4).Value
                ' Diğer sütunları da gerektiği gibi ekleyebilirsiniz
                sonSatir = sonSatir + 1
            Next i
            For i = 8 To 9 'verinin alınacağı ilk satır
                hedefSayfa.Cells(sonSatir, 1).Value = ws.Cells(i, 1).Value
                hedefSayfa.Cells(sonSatir, 2).Value = ws.Cells(i, 2).Value
                hedefSayfa.Cells(sonSatir, 3).Value = ws.Cells(i, 3).Value
                'hedefSayfa.Cells(sonSatir, 4).Value = ws.Cells(i, 4).Value
                ' Diğer sütunları da gerektiği gibi ekleyebilirsiniz
                sonSatir = sonSatir + 1
            Next i
        End If
    Next ws
End Sub
