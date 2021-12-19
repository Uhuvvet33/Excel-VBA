Option Explicit
'Interaction MsgBox
'----------------------------------------------------------

Sub BilgiTopla()
Dim cvp As VbMsgBoxResult   'değişken olarak tanımlaması
    cvp = MsgBox("İşlemler tamam mı?", vbYesNo) 'nitelikli olmayan bilgi toplama mesajı
    If cvp = 6 Then     'yes demek oluyor, (cvp = vbYes) şeklinde kullanılabilir.
        MsgBox "İşlem Tamam"
    Else
        MsgBox "İşleme tamamlayınız. Kolay gelsin." 'bu bilgi mesajı
                'sadece bilgi verilecekse pareantez gerekmez
        Exit Sub    'çıkış yapılır
    End If

End Sub
'----------------------------------------------------------

Sub IptalYok()
Dim hata As Error
Dim cvp As VbMsgBoxResult

On Error GoTo hata
cvp = MsgBox("Cevabınız ?", vbYesNo)
'diğer komutlar
Exit Sub
hata:
Debug.Print Err.Description
End Sub
'----------------------------------------------------------

Sub IptalVar()
Dim cvp As VbMsgBoxResult   'değişken olarak tanımlaması
'Dim hata As Error
'On Error GoTo hata
    cvp = MsgBox("Cevab verir misin?", vbYesNoCancel) 'iptal seçeneğide ekleniyor
    If cvp = vbYes Then
        MsgBox "Evet Seçeneği."
    ElseIf cvp = vbNo Then
        MsgBox "Hayır Seçeneği."
    Else
        MsgBox "İptal Seçeneği."
    End If
'Exit Sub
'hata:
'Debug.Print Err.Description
End Sub
