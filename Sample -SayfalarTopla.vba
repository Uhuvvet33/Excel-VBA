
'-----------------Farklı sayfalardaki verileri sayfa1 de toplamak

Sub AktarTopla()

Dim OgrTpl(39) As Integer
Dim OgrEk(39) As Integer

For j = 2 To Sheets.Count   'Sayfalar Arasında Gezinmek
For i = 1 To 39             'Sütünlarda gezinmek
    OgrTpl(i) = Sheets(1).Range("C6").Cells(i, 1).Value 'Sayfa1 deki değerleri almak
    OgrEk(j) = Sheets(2).Range("C6").Cells(i, 1).Value  'Sayfa2 deki değerleri almak
    Sheets(1).Range("C6").Cells(i, 1).Value = Int(OgrTpl(i)) + Int(OgrEk(j))  
        'Değerleri toplayım Sayfa1 deki hücrelere yazmak
Next i
Next j
  Sheets(1).Range("K3").Value = Format(Now) 'İşlem tarihi ve Saati
  Sheets(1).Select                          'Sayfa1 Seçmek
  
End Sub
