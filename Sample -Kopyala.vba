Sub Kopyala()
  Sheets("Sayfa1").Range("A1:C10").Copy
  Sheets("Sayfa2").Range("A1").PasteSpecial
  Application.CutCopyMode = False
  MsgBox "Kopyalama Yapıldı..!!"
End Sub
'---------------------------------------------------------------
Sub Kopyala2()
  Sheets("Sayfa1").Range("A1:C10").Copy
  sat = Sheets("Sayfa2").Cells(65536, "A").End(xlUp).Row + 5
  Sheets("Sayfa2").Range("A" & sat).PasteSpecial
  Application.CutCopyMode = False
  MsgBox "Kopyalama Yapıldı..!!"
End Sub
