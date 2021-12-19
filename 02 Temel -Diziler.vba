
Sub DiziOrnek1()
'Birden fazla tanımlamalar
Dim bolge1 As String
Dim bolge2 As String
Dim bolge3 As String
'...
Dim bolge20 As String
'Birden fazla tanımlamaları bu şekilde yapmak yerine aşağıdaki şekilde dizi tanımlamak daha verimli olur

Dim bolge(1 To 20) As Integer   'Yukarıdaki şekilde tek tek yerine bu şekilde tanımlanabilir.
End Sub

Sub DiziOrnek2()
'Dizi veya Array değişken tanımlamalar
Dim myDizi1(5) As Integer   '(5) şeklinde üst  sınır belirtilir,
                            'başlangıç değeri 0 olan 6 tane değer alır.
'Başlangıç değerinin 1 den başlamasını istersek, "Option Base 1" şeklinde en üste eklenir.
Dim myDizi2(0 To 5) As Integer  'başlangıç değeride verilerek belirtilir.
Dim myDizi3() As Integer    'Sınırları sonradan belirlenir.
End Sub

Sub DiziOrnek3()
Dim i As Integer
Dim dizi1(1 To 12) As String
For i = 1 To 12
    dizi1(i) = Sheets(1).Cells(i, 1)
Next i
End Sub

Sub DiziIyiAtama()
Dim i As Integer
Dim aylar(1 To 12) As String
Dim segment As Variant

'değişkenlere değer atama ve çağırma
For i = 1 To 12
    aylar(i) = Sheets(1).Cells(i, 1)    '1 hücreden başlayarak 12 değeri aylar dizisine aktarılıyor.
Next i
segment = Array("Kobi", "Bireysel", "Ticaret")  'Variant tipi değişkene üç değer atanıyor
Debug.Print aylar(2)        'aylar dizisinin 2. değeri çağrılıyor
Debug.Print segment(2)      'segment dizisinin 2 değeri çağrılıyor

'Diziden LBound ve UBound kullanarak değerleri çağırmak
Debug.Print "_____________________"
Dim k As Integer
For k = LBound(aylar) To UBound(aylar)  'Bu şekilde aylar dizisinin başlangıç değerini ve bitiş değerine göre
    Debug.Print aylar(k)
Next k

'Diziden sadece değer çağırmak
Debug.Print "---------------------"
Dim ay As Variant
For Each ay In aylar    'For Each döngüsü ile sadece okuma yapılabiliyor.
    Debug.Print ay
Next ay
End Sub
Sub DinamikDizi()
Dim subeler() As String

'..... x hesaplanır
ReDim subeler(x) 'x ile dizinin boyutu sonradan belirleniyor.
'ReDim sadece dinamik dizilerde kullanılıyor.

Dim statikVar As Variant    'Boyut belirtilmez ise dinamik tipli oluyur
Dim dinamikVar(5) As Variant ' statitik tipli oluyor, bunlar üzerinde ReDim işlemi yapılmaz.
End Sub

Sub IkiBoyutluDizi()
Dim mudur(1 To 10, 1 To 2) As Long      'İki boyutlu değişken tanımlama
Dim i As Integer, j As Integer

For i = LBound(mudur, 1) To UBound(mudur, 1)
    For j = LBound(mudur, 2) To UBound(mudur, 2)
        mudur(i, j) = Cells(i + 1, j + 1).Value   'mudur iki boyutlu diziye hücrenin değerleri aktarılıyor
    Next j
Next i

Debug.Print mudur(1, 2) '1 in 2. değeri çağrılıyor
End Sub