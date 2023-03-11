Private Sub ListBoxDoldur()
Application.ScreenUpdating = False
Dim Son As Integer
'Dim ListBxPers As Object

Son = Sheets(4).Cells(Rows.Count, 3).End(3).Row
With frmPersonel.lstBxPers
.ColumnCount = 7
.ColumnHeads = True
.ColumnWidths = "30;100;50;70;70;50;40"
.Clear
.RowSource = "Personel_Sayfasi!A2:G" & Son
End With

Application.ScreenUpdating = True
End Sub
