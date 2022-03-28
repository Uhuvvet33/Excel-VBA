Sub HucreFormulYaz()
' Faklı sayfalardaki c6 hücresini toplama
' Hücreye formul yazdırarak işlem yapma =TOPLA yazıldığı zaman hata veriyor, 
' vba ekranında =SUM yazmak gerekiyor.
  
  Sheets(1).[K11].Formula = "=SUM('" & Sheets(2).Name & ":" & Sheets(Sheets.Count).Name & "'!C6)"
  Sheets(1).[K10].Formula = "=SUM(C6:c20)"
  
' Buradaki kullanımda hesaplayarak işlem yapılıyor. 
' Hücreye formül yazılmıyor.
  Sheets(1).[K11].Value = WorksheetFunction.Sum(Range("c6:c20"))
End Sub

Sub HucreTopla()
With sheets(2).Range("W7:W70")
.Formula = "=IFERROR(IF(SUM(E7:V7)>0,SUM(E7:V7),""""),"""")"
.Value = .Value
End With
End Sub
