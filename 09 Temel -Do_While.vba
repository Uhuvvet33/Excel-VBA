Sub Do_While1()
Do  'En az bir kez çalışacaktır.
    sayı = InputBox("Karesi alınacak bir sayı girin")
Loop While Not IsNumeric(sayı)  'Girilen sayı bir rakam mı?
MsgBox sayı & " sayısının karesi şudur:" & sayı * sayı
End Sub
'----------------------------------------------------

Sub Do_While2()
Do While Not IsNumeric(sayı) 'değer False olduğu için döngüye girmez.
    sayı = InputBox("Karesi alınacak bir sayı girin")
Loop
MsgBox sayı & " sayısının karesi şudur:" & sayı * sayı
End Sub
'----------------------------------------------------

Sub DO_Until()
Application.DisplayAlerts = False 'Herhangi bir uyarı çıkartma.
If Sheets.Count < 5 Then          'Sayfa sayısı 5 ten küçükse
    Do Until Sheets.Count = 5
        Sheets.Add After:=Sheets(Sheets.Count) 'Sayfa Sayısı ekler
    Loop
ElseIf Sheets.Count > 5 Then 'Sayfa sayısı 5'ten büyükse
    Do Until Sheets.Count = 5
        Sheets(Sheets.Count).Delete  'Sayfa sil
    Loop
End If
End Sub
'----------------------------------------------------

Sub Do_Until2()
Do
    Sheets(2).Delete 'Her defasında hep 2.sayfa silinir, ta ki tek sayafana kalana kadar
Loop Until Sheets.Count = 1
End Sub
'----------------------------------------------------

Sub Dongu_Exit()
i = 1
Do While i < 1000
    If IsEmpty(Cells(i, 1)) Then
        Exit Do
    End If
    i = i + 1
Loop
End Sub
'----------------------------------------------------

Sub MaxIf()
'maxı yazacağın yere gel, ordayken çalıştır ve liste sıralı olsun

Set kriter = Application.InputBox("ana değişken kriterinin olduğu sütundan bir hücre seç", Type:=8)
Set rakam = Application.InputBox("maksimumun arandığı sütunu seç", Type:=8)

ks = ActiveCell.Column - kriter.Column
rs = ActiveCell.Column - rakam.Column

Do
    Maks = ActiveCell.Offset(0, -rs).Value
    Set ilkyer = ActiveCell
    
    Do While ActiveCell.Offset(0, -ks).Value = ActiveCell.Offset(1, -ks).Value        
        If ActiveCell.Offset(1, -rs).Value > Maks Then Maks = ActiveCell.Offset(1, -rs)
        ActiveCell.Offset(1, 0).Select                
    Loop
    
    Set sonyer = ActiveCell
    Range(ilkyer, sonyer).Value = Maks
    ActiveCell.Offset(1, 0).Select
        
Loop Until ActiveCell.Offset(0, -1).Value = ""
End Sub
