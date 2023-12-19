Sub sheet_vs_worksheet()
    Debug.Print ActiveWorkbook.Sheets.Count
    Debug.Print ActiveWorkbook.Worksheets.Count
End Sub

Sub erisim()
    Worksheets.Item(1).Select
    Worksheets(1).Select

    Sheet4.Select 'kod isimle. (sadece bunda intelisense çıkar, ilgili wb içinde)

    Sheets(1).Select 'Sheets koleksiyonu ve index
    Worksheets(1).Select 'Worksheets koleksiyonu ve index
    Sheets("Sheet3").Select 'Sheets koleksiyonu ve sayfa adı
    Worksheets("Sheet2").Select 'Worksheets koleksiyonu ve sayfa adı
End Sub

Sub activesheet_deklerasyon_intelisense()
    Dim ws As Worksheet
    Set ws = ActiveSheet
End Sub

Sub activate_ve_gezinme()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        Range("A1").Value = i
        i = i + 1
    Next

    'recoreder, ctrl+pgup/pddown
    ActiveSheet.Next.Select ' sonraki sayfa
    ActiveSheet.Previous.Select ' önceki sayfa
End Sub

Sub gizlisec_activateli()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select ' ws.Select gizli olduğu için hata verir
        Range("a1") = ws.Index
    Next
End Sub

Sub gizliyisec2()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible Then
            Range("a1") = ws.Index
            ws.Select
        End If
    Next
End Sub

Sub isim()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Name = "Yeni sayfa" 'yaz
    MsgBox ws.Name 'oku
End Sub

Sub ekle_sil_gizle_taşı_kopy()
    'recoder
    Set yeni = Worksheets.Add
    Worksheets.Add
End Sub

Sub sayfakoruma()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ProtectContents = True Then
    ws.Unprotect Password:="1234"
    End If

    On Error GoTo hatayakala

    Range("A1") = Environ("USERNAME")
    'çeşitli işlemler
    '0'a bölme olacak ve hata alacak
    a = InputBox("Bir sayı girin")
    b = InputBox("Bu sayıyı kaça bölelim")
    MsgBox "Sonuç: " & a / b

    ws.Protect Password:="1234"
    Exit Sub

    hatayakala:
    ws.Protect Password:="1234"
End Sub

Sub sort_filter()
    'sort:recorder-->range resize örneği
    Set alan = Range("A1").CurrentRegion 'başlık dahil

    Range("A2").Select
    ActiveWorkbook.Worksheets("sıralama").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sıralama").Sort.SortFields.Add Key:=Range( _
    "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal

    With ActiveWorkbook.Worksheets("sıralama").Sort
        .SetRange alan
        .Header = xlYes 'noyu yes yaptık
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'filter-->recorder
    'başlık satırı dahildir. Resize ve Offset gereksiz
        Selection.AutoFilter 'recordeela gelir, manuel kodda gerek yoktur
        ActiveSheet.Range("$A$1:$B$7").AutoFilter Field:=2, Criteria1:="<130", _
            Operator:=xlAnd 'recordeela gelir, manuel kodda gerek yoktur
        ActiveSheet.Range("$A$1").CurrentRegion.AutoFilter Field:=2, Criteria1:="<130"
        ActiveSheet.ShowAllData
        Selection.AutoFilter
End Sub

Sub filtresub()    
    a = worksheetfunction.RandBetween(1, 4)
    MsgBox "Case " & a & " gerçekleşecek"
    
    Select Case a
        Case 1 'o an kapalıysa kapalı kalmaya devam, açıksa kapanır
            ActiveSheet.AutoFilterMode = False 'buna sadece false atanıyor, true atanamaz
        Case 2 'açıkken filtre konursa kapanır
            If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
        Case 3 'kapalıyken filtre konursa açılır
            If ActiveSheet.AutoFilterMode = False Then Selection.AutoFilter
        Case 4 'filtre okları aktif olsa da olmasa da burası çalışır
            ActiveSheet.Range("$A$1").CurrentRegion.AutoFilter Field:=2, Criteria1:=Range("b2")
    End Select
    
    If ActiveSheet.AutoFilterMode = True Then 'filtre okları açıksa. Case 3 veya 4
        If ActiveSheet.FilterMode = False Then
            MsgBox "Case 3:Filtre açık ama kriter yok"
        Else
            MsgBox "Case 4:filtre açık ve kriter var, şimdi kriter kaldırılacak, ama filtre açık kalacak"
            ActiveSheet.ShowAllData
        End If
    Else 'filtre okları yoksa, yani henüz bir autofilter düğmesine basılmamışsa
        MsgBox "Case 1 veya 2:filtre yok"
    End If
End Sub

Sub pastespecial()
    ActiveChart.ChartArea.Copy
    Range("M2").Select
    ActiveSheet.pastespecial Format:="Picture (PNG)", Link:=False, _
    DisplayAsIcon:=False
End Sub

Sub parenti()
    Debug.Print TypeName(ActiveSheet.Parent)
    Debug.Print ActiveSheet.Parent.Name
End Sub
