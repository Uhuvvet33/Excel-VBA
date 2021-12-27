'Select Case
Sub Slct()
Dim sayi As Integer
sayi = InputBox("Sayı giriniz :")

Select Case sayi
    Case Is < 0
        MsgBox "Negatif"
    Case Is > 0         ' sıfır girildiği zamanilk bunu gördüğü için bu mesaj gelir
        MsgBox "Pozitif"
    Case 1 To 9         ' 0 ile 9 arasında
        MsgBox "1 ile 9 arası rakamlar"
    Case 0              ' en çok kullanılan
        MsgBox "sıfır"
    Case Else
        MsgBox "Lütfen sayı giriniz."
End Select
End Sub
'--------------------------------------------------

Sub ChooseOrnek()
Dim ay As Integer
Dim ayad 'as String

ay = Application.InputBox("Ay Numarası Giriniz", Type:=1)
ayad = Choose(ay, "Ocak", "Şubat", "Mart")

Debug.Print ayad   'Çıktı : Şubat
End Sub
'--------------------------------------------------

Sub SwitcOrnegi()
KanalKodu = InputBox("Kanal Kodu :")
KanalAdi = Switch(KanalKodu = 1, "Şube", KanalKodu = 8, "İnternet")

Debug.Print KanalAdi
End Sub
