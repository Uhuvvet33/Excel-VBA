
'-----------------Farklı sayfalardaki verileri sayfa1 de toplamak

Sub VeriTpl()
  
'Bilgileri Toplama
Dim OgrTpl As Integer
Dim OgrEk As Integer

Dim KatTpl As Integer
Dim KatEk As Integer

Dim AyrTpl As Integer
Dim AyrEk As Integer

For j = 2 To Sheets.Count 'Sayfalar Arasında Gezinmek

For i = 1 To 18
    OgrTpl = Sheets(1).Range("C6").Cells(i, 1).Value
    OgrEk = Sheets(j).Range("C6").Cells(i, 1).Value
    Sheets(1).Range("C6").Cells(i, 1).Value = Int(OgrTpl) + Int(OgrEk)
    
    KatTpl = Sheets(1).Range("E6").Cells(i, 1).Value
    KatEk = Sheets(j).Range("E6").Cells(i, 1).Value
    Sheets(1).Range("E6").Cells(i, 1).Value = Int(KatTpl) + Int(KatEk)

    AyrTpl = Sheets(1).Range("F6").Cells(i, 1).Value
    AyrEk = Sheets(j).Range("F6").Cells(i, 1).Value
    Sheets(1).Range("F6").Cells(i, 1).Value = Int(AyrTpl) + Int(AyrEk)
Next i

Next j
