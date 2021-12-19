'IF ELSEIF ENDIF; ÖRNEKLERİ

Sub ifornegi1()
'boolen örneği
Dim ilkdeger As Boolean 'değişken tanımlama

isim = InputBox("İsminizi giriniz :")
If isim <> "" Then
    For i = 1 To 5
        If ilkdeger = False Then    'ilkdeger False ise
            kelime = isim
        Else
            kelime = kelime + "; " + isim
        End If
    ilkdeger = True             'ilkdeger for sonunda deger atıyoruz
    Next i
    MsgBox kelime       'değerleri gösterelim
Else
    MsgBox "Bilgi girmediniz ..."
    Exit Sub
End If
End Sub

Sub ifornegi2()
'boolen örneği, ilk değer olmadan işlem yapmak
Dim ilkdeger As Boolean 'değişken tanımlama

isim = InputBox("İsminizi giriniz :")
If isim <> "" Then
    kelime = isim
    For i = 1 To 5
        kelime = kelime + "; " + isim
    Next i
    MsgBox kelime       'değerleri gösterelim
Else
    MsgBox "Bilgi girmediniz ..."
    Exit Sub
End If
End Sub

'MANTIKSAL SORGULAMALAR

Sub Mantıksal()
If IsNumeric(3) Then Debug.Print "Bu bir sayıdır"      'Varsayılan True
If IsEmpty(ActiveCell) Then Debug.Print "Bu boş bir hücredir"
If Not IsEmpty(ActiveCell) Then Debug.Print "Bu boş değildir"   'Not ile kullanım
If IsDate(Date) Then Debug.Print "Bu bir tarihtir"
deger = 4
If IsNull(deger) Then Debug.Print "Bu bir Null'dur"
'Empty ve Null farklılıkları vardır
End Sub

'If - ElseIf - EndIf ;  DETAYLARI

Sub ifnot1()
Set alan = ActiveCell       'aktif hücreyi alan değişkeniden aktarıyoruz
If Not IsEmpty(alan) Then
    'hücre boş değilse yapılacak işlemler
    MsgBox alan
Else
    MsgBox "Boş bir hücre seçtiniz."
End If
End Sub

Sub ifnot2()
If Not DosyaVarmi(dosyaadi) Then
    Exit Sub    'dosya yoksa çık
Else
    'dosya var ve işlemleri yap
End If
End Sub

Sub IfNot_IsNothing()
Dim hucre As Range
Dim ws As Worksheet
Dim wb As Workbook

Set hucre = ActiveCell
Set ws = ActiveSheet

If hucre Is Nothing Then    'hucre değişkenine değer atanım atanmadığı, boş olup olmadığı IsEmpty ile sorgulanır
    'hucre değişkenine atama yapılmamış
Else
    'atama yapılmış
End If

If Not ws Is Nothing Then 'birşey ise
    'atama yapılmış
Else
    'yapılmamış
End If

'sadece birşey olma durumu, en çok kullanım şekillerinden
If Not wb Is Nothing Then
    'değer atanmamış ise işlem yapmayacak
End If
End Sub
