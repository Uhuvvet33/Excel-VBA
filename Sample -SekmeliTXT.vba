Sub txt_BRN()
Set s1 = Sheets("AktarilacakSayfa"): s1.Activate
yol = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator
adı = InputBox("", "TXT Belge için İSİM yazınız:")
If adı = "" Then Exit Sub
If Dir(yol & adı & ".txt") <> "" Then
    MsgBox "Bu dosya mevcuttur."
    Exit Sub
End If
Open yol & adı & ".txt" For Output As #1
    For i = 2 To s1.[A65000].End(3).Row
        Print #1, Cells(i, 3), Cells(i, 4), Cells(i, 5), Cells(i, 6), Cells(i, 7), Cells(i, 9) 'Sütunlar burdan eklenebilir
    Next i
Close #1
MsgBox "Masaüstü'ne " & adı & " adlı txt blge kaydedildi..", , "Ö. BARAN'a teşekkürler"
End Sub
