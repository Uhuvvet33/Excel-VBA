Sub Operatörler()
Dim i As Integer
i = 0
i = 8 \ 3   'mod alma, 2
i = 8 Mod 3 'mod alma, 2
i = 3 ^ 2   'üst alma, 9
End Sub

Sub BoolDegerleri()
Dim a As Boolean
Dim x, y

a = True    'a ya boolen değerlerinden True yapılıyor
If a And (x = 0 Or y = 1) Then 'if a=True demek yerine
    MsgBox "Doğru"
Else
    MsgBox "yanlış"
End If

'a = Not a      'a nın değeri True ise False olsun diyoruz
End Sub

Sub Birlestirme()
Dim a As String, b As String, c As Integer, d As Integer, e As String
a = "10"
b = "20"
c = 300
d = 30
e = "Mehmet"

Debug.Print "Merhaba " + e   'değerler string ise birleştirir, integer ise toplar
Debug.Print "Merhaba " & e   'her zaman birleştirir
Debug.Print c & d            'birleştirir (& her zaman birleştirir
Debug.Print c + d            'toplar

Dim mesaj, goster
mesaj = "Merhaba, Dünya" & vbCrLf   '"vbCrLF"(enter) veya "Chr(10)" bir alt satıra geçmek
mesaj = mesaj + "Ay" & Chr(10) & Chr(10)
mesaj = mesaj & "Yıldızlar"

goster = InputBox(mesaj)
End Sub

Sub SayiArttir()
Dim i As Integer
i = 0
Do              'döngü
    i = i + 1
    Debug.Print i
Loop Until i = 100
End Sub
