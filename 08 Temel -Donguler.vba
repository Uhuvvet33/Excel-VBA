'DÖNGÜLER

Sub for1()
For i = 1 To 10         '1 den 10 kadar döngü
    Cells(i, 1) = i     'İlk hücreden itibaren artarak alta yazıyor
Next i

For i = 10 To 1 Step -2     '-2 azaltarak
    Cells(i / 2, 1) = i + 1 '5 satırdan itibaren yukarı doğru
Next i

For i = 1 To Sheets.Count   'Sayfa sayısı
    Worksheets(i).Select    'Sayfalar arasında gezmek
Next i
End Sub
'----------------------------------------------------

Sub IcIceDongu()
For i = 1 To 5          '5 sütun
    For k = 1 To 11     '11 satır
        Cells(k, i).Select  'hücreleri seçerek gezdiriyoruz.
        If IsEmpty(Cells(k, i)) Then 'Hücre boş ise
            Cells(k, i).Interior.Color = vbYellow 'hücre nin rengi
            adet = adet + 1     'boş hücreleri saydırıyoruz
        End If
    Next k
Next i
MsgBox "Toplamda " & adet & " adet boş hücre."
End Sub
'----------------------------------------------------

Sub IcIceDongu2()
'i ve j'yi kullanmayacaksam
Dim adet As Integer
Dim hucre As Range, alan As Range

Set alan = Range("A1").CurrentRegion    'Alan Seçme Ctrl+a

For Each hucre In alan        'Seçilen hücrelerde dolaş
    If IsEmpty(hucre) Then
        hucre.Interior.Color = vbYellow
        adet = adet + 1
    End If
Next hucre

MsgBox "toplamda " & adet & " adet boş hücre var"
End Sub
'----------------------------------------------------

Sub DonguExit()
Dim a As String
Dim hucre As Range, alan As Range

Set alan = Range("A1").CurrentRegion

For Each hucre In alan
    hucre.Select
    If IsEmpty(hucre) Then
        a = hucre.Address
        Exit For
    End If
Next hucre

If a <> vbNullString Then
    MsgBox "ilk olarak " & a & " adresinde boşluğa rastlanmıştır"
Else
    MsgBox "herhangi bir boş hücre bulunmamaktadır"
End If
End Sub
