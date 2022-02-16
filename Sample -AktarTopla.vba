Sub AktarTopla()

Dim OgrTpl(39) As Integer
Dim OgrEk(39) As Integer

Dim KatTpl(39) As Integer
Dim KatEk(39) As Integer

Dim AyrTpl(39) As Integer
Dim AyrEk(39) As Integer

Dim OgrTpl2(39) As Integer
Dim OgrEk2(39) As Integer

Dim KatTpl2(39) As Integer
Dim KatEk2(39) As Integer

Dim AyrTpl2(39) As Integer
Dim AyrEk2(39) As Integer

Dim OgrTpl3(39) As Integer
Dim OgrEk3(39) As Integer

Dim KatTpl3(39) As Integer
Dim KatEk3(39) As Integer

Dim AyrTpl3(39) As Integer
Dim AyrEk3(39) As Integer

Application.ScreenUpdating = False  'Ekran titremesini iptal

For i = 1 To 39
    OgrTpl(i) = Sheets(1).Range("C6").Cells(i, 1).Value
    OgrEk(i) = Sheets(2).Range("C6").Cells(i, 1).Value
    Sheets(1).Range("C6").Cells(i, 1).Value = Int(OgrTpl(i)) + Int(OgrEk(i))
    
    KatTpl(i) = Sheets(1).Range("E6").Cells(i, 1).Value
    KatEk(i) = Sheets(2).Range("E6").Cells(i, 1).Value
    Sheets(1).Range("E6").Cells(i, 1).Value = Int(KatTpl(i)) + Int(KatEk(i))

    AyrTpl(i) = Sheets(1).Range("F6").Cells(i, 1).Value
    AyrEk(i) = Sheets(2).Range("F6").Cells(i, 1).Value
    Sheets(1).Range("F6").Cells(i, 1).Value = Int(AyrTpl(i)) + Int(AyrEk(i))
'---
    OgrTpl2(i) = Sheets(1).Range("C50").Cells(i, 1).Value
    OgrEk2(i) = Sheets(2).Range("C50").Cells(i, 1).Value
    Sheets(1).Range("C50").Cells(i, 1).Value = Int(OgrTpl2(i)) + Int(OgrEk2(i))
    
    KatTpl2(i) = Sheets(1).Range("E50").Cells(i, 1).Value
    KatEk2(i) = Sheets(2).Range("E50").Cells(i, 1).Value
    Sheets(1).Range("E50").Cells(i, 1).Value = Int(KatTpl2(i)) + Int(KatEk2(i))

    AyrTpl2(i) = Sheets(1).Range("F50").Cells(i, 1).Value
    AyrEk2(i) = Sheets(2).Range("F50").Cells(i, 1).Value
    Sheets(1).Range("F50").Cells(i, 1).Value = Int(AyrTpl2(i)) + Int(AyrEk2(i))
'---
    OgrTpl3(i) = Sheets(1).Range("C94").Cells(i, 1).Value
    OgrEk3(i) = Sheets(2).Range("C94").Cells(i, 1).Value
    Sheets(1).Range("C94").Cells(i, 1).Value = Int(OgrTpl3(i)) + Int(OgrEk3(i))
    
    KatTpl3(i) = Sheets(1).Range("E94").Cells(i, 1).Value
    KatEk3(i) = Sheets(2).Range("E94").Cells(i, 1).Value
    Sheets(1).Range("E94").Cells(i, 1).Value = Int(KatTpl3(i)) + Int(KatEk3(i))

    AyrTpl3(i) = Sheets(1).Range("F94").Cells(i, 1).Value
    AyrEk3(i) = Sheets(2).Range("F94").Cells(i, 1).Value
    Sheets(1).Range("F94").Cells(i, 1).Value = Int(AyrTpl3(i)) + Int(AyrEk3(i))
Next i

Sheets(1).Range("K3").Value = Format(Now, "dd.mm.yyyy hh:mm")
End Sub
