Function VirgulEk(Veri As String) As String
' hücre içindeki değerin aralarına virgün ekler,
' veri kaç karakter olduğu önemli değil
' Fonksiyon olarak çağırırken a2 hücresine  =VirgulEk(A1)  yazılması yeterli.

  Dim i As Integer
    Dim krk, str, deger As String
    str = Veri
    krk = Len(str)
    For i = 1 To Len(str)
        deger = deger & Mid(str, i, 1) & ","
        ' Virgül yerine başka bir karakter eklenebilir
    Next i
    VirgulEk = Mid(deger, 1, Len(deger) - 1)

End Function
