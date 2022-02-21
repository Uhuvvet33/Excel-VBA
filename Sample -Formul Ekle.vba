Sub HucreFormulYaz()
  ' Faklı sayfalardaki c6 hücresini toplama
  Sheets(1).[K11].Formula = "=SUM('" & Sheets(2).Name & ":" & Sheets(Sheets.Count).Name & "'!C6)"

End Sub
