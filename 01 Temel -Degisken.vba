Option Explicit    'Değişkenleri tanımlanmanızı ister, zorlar.
'----------------------------------------------------------
Sub YorumEklemek()
'Burda olduğu gibi yorum satırı tek tırnak şeklinde
Rem -başına rem yazdığımız zamanda yorum satırı olabiliyor

Range("A1").Value = "Mehmet"        '"A1" hücresine değer yazdırılır.
Range("A1").Interior.Color = vbRed  'Hücrenin rengini değiştirilir
Range("A1").Clear                   'Hücre temizleniyor

Range("TanimAd").Value = "Mustafa"          ' Sayfa üzerinden tanımlanan hücreye değer ataması
Range("TanimAd").Interior.Color = vbBlue    'Tanımlı Hücrenin rengini değiştirilir
Range("TanimAd").Clear                      'Tanımlı Hücre temizleniyor

End Sub
'----------------------------------------------------------

Sub DegiskenTanimlamak()
Const bolgeSayisi As Integer = 10   'Sabit değer tanımlaması
Dim isim As String                  'Karakter değişken tipi
Dim mini As Byte                    '0 ile 255 arası değer alabilir.
Dim sayi As Integer                 'Sayı veri tipi örnek : 3
Dim uzunSayi As Long
Dim kisaSayi As Single
Dim agirlik As Double       'Kayan veri tipi örnek : 3,17
Dim gerceklesme As Boolean  'True ve False değerlerini alır.
Dim tarih As Date           'Tarih
Dim cesitli As Object       'Çok çişitli
Dim hersey As Variant       'Varsayılan veri tipi(herşey olablir)
Dim hersey2                 'Bu şekilde varyant tanımlama oluyor
Dim sabitBoy As String * 15 'Sabit boyutlu örneğin "Mehmet         " uzunluğu 15 olana kadar sonuna boşluk ekler
Dim r1, r2 As Integer       'Bu şekilde tanımlamada r1 variant olarak tanımlanır.
Dim r1 As Integer, r2 As Integer

Dim wb As Workbook          'Nesne tanımlama örneği
Set wb = ActiveWorkbook     'Nesne tanımlamada set kullanılmalıdır.
End Sub
