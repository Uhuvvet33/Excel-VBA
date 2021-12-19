'Interaction InputBox;  sadece giriş yapılır veilerlenir.
'klasik InputBox kullanımı

Sub Giris1()
a = InputBox("Bir sayı giriniz:")
b = InputBox("İkinci bir sayı giriniz:")

Range("B1").Value = a + b
End Sub
Sub Giris2()
a = InputBox("Bir sayı giriniz:")
b = InputBox("İkinci bir sayı giriniz:")

Range("B1").Value = Val(a) + Val(b)  'Val ile integer dönüşümü yapılıyor
End Sub

Sub Giris3()
Dim a As Integer, b As Integer
a = InputBox("Bir sayı giriniz:")
b = InputBox("İkinci bir sayı giriniz:")

Range("B1").Value = a + b  'değişkenler Dİm ile integer olarak tanımlandı
End Sub

Sub InputSyntax() 'InputBox  syntax
a = InputBox("il kodu girin", "il kodu", 34)
             'verilen mesaj, bar üstünde, varsayılan 34 değeri
             
End Sub

'Aplication InputBox; Aplication class altında yer alır, farkı sonuda type vardır.

Sub ApplicationInputbox()
Dim hucre As Range
Set hucre = Application.InputBox(prompt:="Son hücreyi seçin", Type:=8)
'En çok kullanılan typ türleri 1;Sayı , 2;Metin , 8;Range
MsgBox hucre.Address    'seçilen hücreyi gösteriyoruz.
End Sub


' InputBox ; DETAYLAR
'----------------------------------------------------------------
Sub InputIptal2()
'klasik InputBox
deger = InputBox("bir değer giriniz:")      'varsayılan degeri String'tir
If deger <> "" Then
    MsgBox "Girilen değer :" & deger
Else
    MsgBox "Bir değer girmediniz ..."
End If
End Sub

Sub InputIptal3()
'Application'lı, String değer için iptal kontrolü
Dim a As String
a = Application.InputBox("Adınızı girin :", Type:=2)    'iptal edildiğinde dönüş değeri False
If a <> "False" And a <> "" Then
    'False string olduğu için tırnak içerisine alınır. Boş değer girildiğinide aynı anda kontrol ediyoruz.
    MsgBox "Girilen değer: " + a
Else
    MsgBox "Değer girmediniz ..."
End If
End Sub

Sub InputIptal4()
'Application'lı, Integer değer için iptal kontrolü
Dim a As Integer
a = Application.InputBox("Adınızı girin :", Type:=1)  'type değeride :=1 yapılıyor.
'Type değeri sayı olduğuiçin 1 yapılıyor, dönüş değeri iptal olduğu zaman 0 veya False
If a <> 0 Then 'False yerine 0 yazılabilir
    'False Integer olduğu için tırnak içerisine alınmaz.
    MsgBox "Girilen değer: " & a
Else
    MsgBox "Değer girmediniz ..."
End If
End Sub

Sub InputRange()
'Applicationlu, Range
'Range seçiminde eğer kullanıcı seçim yapmazsa hata oluşur, bu yüzden bir hata kontrol mekanizması da ekleriz
've ayrıca bir seçim yapıp yapmadığını da Nothing ile kontrol ederiz
On Error Resume Next 'burayı yazmassak hata alırız. Hata yönetim mekanizmaları için ilgili sayfaya gidip bilgi edinebilirsiniz
Dim a As Range  'hücre seçimlerinde Range kullanıır.
Set a = Application.InputBox("Bir hücre seçin", Type:=8)
If Not a Is Nothing Then
    MsgBox "Seçim yapıldı"
    'Diğer kodlar buraya
Else
    MsgBox "Bir seçim yapılmadan çıkmayı tercih ettiniz"
End If
End Sub

Sub SayfaEkle()
Dim i As Integer, syf As Integer   'değişken tanımlıyoruz
syf = Application.InputBox("Eklenecek sayfa sayısı giriniz: ", Default:=3, Type:=1) 'default:=3, varsayılan değer olarak 3
If syf = False Then     'iptal edilirse
    Exit Sub            'fonksiyondan çıkış
Else
    For i = 1 To syf    '1 den girilen değere kadar
        Worksheets.Add  'sayfa ekleme işlemi
    Next i
End If
End Sub

Sub KullaniciDostuKodlama()
'kullanıcı açısından iyi, kodlamacı açısından da iyi
a = InputBox("Müşteri segmenti için bir değer giriniz. " & vbCrLf & _
"Bireysel müşteriler için 1," & vbCrLf & _
"Ticari müşteriler için 2," & vbCrLf & _
"Kurumsal müşteriler için 3")
'vbCRLF; enter bir alt satırda göstermek
'_ kullanımı, kodların bir alt satırında devam ettiğidir.
End Sub
