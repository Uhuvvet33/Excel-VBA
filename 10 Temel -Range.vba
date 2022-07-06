Sub Secimler()     
'Hücre ve Alan  Seçimleri
Range("A1").Select
Range("A1:B2").Select
Range("A1:B2,C3:D4").Select
Range("Ozelalan").Select    'İsim verilen hücrenin seçilmesi
Range("Ozelalan,ozelalan2").Select
Range(Range("A1"), Range("C5")).Select

Cells(3, 2).Select      'Satır, Sütun şeklinde
Cells(3, "B").Select
[A1].Select             'Hücre seçimin kısa hali
[A1:B5].Select
[Ozelalan].Select

'Sheets(2).Range("A10").Select  'Hatalı birişlem
Application.Goto Sheets(2).Range("A10") 'Belirtilen sayfaya giderek seçim yapmak
Sheets(2).Select            'Sayfa Seçimi
Sheets("Sayfa 2").Select    'Sayfa Seçimi
    Range("A10").Select
End Sub
'----------------------------------------------------

Sub Secim_Maliyeti()
'Select işlemleri zaman kaybına neden olabilir.
'En az sayıda kullanmaya çalışılmalı.
Dim bas As Single, bitis As Single
bas = Timer
For i = 1 To 1000
    'Cells(i, 1).Select   'bu varken 43 sn, yokken 1 sn
    Cells(i, 1) = i
Next i
bitis = Timer
Debug.Print bitis - bas
End Sub
'----------------------------------------------------

Sub satır_sütun_secim()
'sütunlar
Range("A:A").Select
Columns("A").Select
Columns(1).Select
Range("A:C").Select
Columns("A:C").Select
Range("A:C,E:E,H:K").Select
Columns.Select      'Tüm Sütunları  Seçmek İçin
Range("A1").EntireColumn.Select     'Bulunduğumuz tüm kolanları seçmek

'satırlar
Range("1:1").Select
Rows(1).Select
Range("1:5").Select
Rows("1:5").Select
Range("1:5,8:10").Select
Rows.Select     'Tüm satırları seçmek için
Range("A1").EntireRow.Select
End Sub
'----------------------------------------------------

Sub Aktivate_Select()
Range("A1:B10").Select
Debug.Print ActiveCell.Address 'sol üstteki ilk hücre, yani A1
Debug.Print Selection.Address 'tüm alan, yani A1:b10

Range("B8").Activate
Debug.Print ActiveCell.Address 'B8
Debug.Print Selection.Address 'seçim hala aynı, değişmedi, a1:b10

Range("C8").Activate 'ilk seçim alanının dışında bir hücre seçiliyor, artık selection da değişmiştir
Debug.Print ActiveCell.Address 'C8
Debug.Print Selection.Address 'C8
End Sub
'----------------------------------------------------

Sub Ozel_Secimler()
ActiveCell.CurrentRegion.Select
ActiveSheet.UsedRange.Interior.Color = vbYellow 'UsedRange işlem yapılmış bütün hücreleri seçip, sarı yapacak
ActiveCell.CurrentRegion.Cells.SpecialCells(xlCellTypeVisible).Select 'SpecialCells görünenleri seç
End Sub
'----------------------------------------------------

Sub ucnoktalar()
Cells.SpecialCells(xlCellTypeVisible)(1).Select 'Ctrl+Home
Cells.SpecialCells(xlCellTypeLastCell).Select 'ctrl+End
ActiveSheet.UsedRange.Select 'A1'deyken Ctrl+Shift+End

ActiveCell.End(xlUp).Select         'Aktif hücrenin Dolu olan en üste gitmek
ActiveCell.End(xlDown).Select       'Alta
ActiveCell.End(xlToRight).Select    'En Sağa
ActiveCell.End(xlToLeft).Select     'En Sola
End Sub
'----------------------------------------------------

Sub valuetext()
'okuma
    With Range("A1")
        Debug.Print .Text   '21 Ocak 1979 Pazar
        Debug.Print .Value  '21 Ocak 1979
        Debug.Print .Value2 '28876
    End With
'yazma
Range("A2").Value = 1       'Tavsiye edilen kullanım
Range("A3") = 1
End Sub
'----------------------------------------------------

Sub cutcopypaste()
    Range("A1:B5").Select
    Selection.Cut
    Sheets(3).Select    'seçili olan hücreye
    ActiveSheet.Paste   'sheet metodu
End Sub
'----------------------------------------------------

Sub cutcopypaste2()
'başka sayfaya value
    Sheets(1).Select
    Range("A1:B5").Select
    Selection.Copy
    Sheets(3).Select
    ActiveCell.PasteSpecial xlPasteValues 'range metou
End Sub
'----------------------------------------------------

Sub copy2()
    'PasteSpecial xlPasteValues alternatifi, sadece value için
    Range("A1:B5").Copy Sheets(3).Range("A15")
End Sub
'----------------------------------------------------

Sub dogrudanyapıştır()
    Range("A2").Value = Range("A1").Value
End Sub
'----------------------------------------------------

Sub adres()
Debug.Print ActiveCell.Address
Debug.Print ActiveCell.Address(0, 0)  '$'sız yazımı
If ActiveCell.Address = "$A$1" Then Exit Sub 'target kullanımı
End Sub
'----------------------------------------------------

Sub Item_Cells()
    Range("B3:D6").Item(1, 2).Select 'C3 hücresi seçilir
    Range("B3:D6").Item(1).Select 'sol üstteki ilk hücre yani B3 seçilir
    Range("B3:D6").Item(2).Select 'C3 hücresi
    Range("B3:D6").Item(5).Select 'C4 hücresi
    
    'shortcut
    Range("B3:D6")(1).Select
    Cells(2, 3).Select '=Cells.Item(2,3).Select demektir.
End Sub
'----------------------------------------------------

Sub görelisecim()
Dim alan As Range
Set alan = Range("C5:E8")

'item, collectionın üyesidir
alan.Select ' tamamı
alan.Range("A1").Select '???
alan.Item(0, 0).Select ' bir satır sol üst
alan.Item(1, 1).Select 'sol üstteki ilk hücre
alan.Item(1).Select 'sol üstteki ilk hücre
alan.Item(0).Select 'sol üstteki ilk hücrenin bir solu
alan.Cells(1, 1).Select 'sol üstteki ilk hücre
alan.Cells.Select 'tamamı
alan(1, 1).Select 'item gibi davranır
alan(0, 0).Select 'item gibi davarnır
alan(1).Select 'item gibi davranır
alan(0).Select 'item gibi davranır

Range("A1")(5).Select
End Sub
'----------------------------------------------------

Sub offsetornek()
Range("C2").Offset(1, 0).Select     'C3  "Offset kaydırma işlemiyapar."
Range("C2").Offset(-1, 2).Select    'E1
Range("C2").Offset(0, -2).Select    'A2
Range("C2").Offset(0, 0).Select     'C2

Range("C2").EntireRow.Offset(1).Select '3.satır. Range("C2").EntireRow.Offset(1,0).Select ile aynıdır. İkinci parametre yoksa 0 anlamındadır
Range("C2").EntireRow.Offset(-1).Select '1.satır
Range("C2").EntireColumn.Offset(, -1).Select '2.sütun. Range("C2").EntireColumn.Offset(0, -1).Select ile aynıdır. ilk paramterde 0 yerine boş da geçilebilir

Range("C2:F6").Offset(1, 1).Select 'D3:G7 seçilir
End Sub
'----------------------------------------------------

Sub resizeornek1()
Range("C3:G7").Select
'Resize belli bi alanı yniden boyutlandır.
Selection.Resize(Selection.Rows.Count - 1, Selection.Columns.Count + 2).Select

Range("C3:G7").Resize.Select        ' aynen kalır
Range("C3:G7").Resize().Select      ' aynen kalır
Range("C3:G7").Resize(1).Select     'kolon parametresi yok, o yüzden aynı kalır, satır ise 1 satır olacak şekilde daralır
Range("C3:G7").Resize(, 2).Select   'satır parametresi yok, o yüzden aynı kalır, sütun ise 2 sütun olacak şekilde daralır
End Sub
'----------------------------------------------------

Sub resizeiledinamiksıralama()
    enalt = Range("A1").End(xlDown).Row
    Set alan = Range("A1").CurrentRegion.Resize(enalt - 1)
    alan.Select
    Set alan = alan.Offset(1)
    alan.Select
    'Set alan = Range("A1").CurrentRegion.Resize(enalt - 1).Offset(1)
    
    Range("A2").Select
    ActiveWorkbook.Worksheets("Sayfa6").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sayfa6").Sort.SortFields.Add Key:=Range( _
    "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
    
    With ActiveWorkbook.Worksheets("Sayfa6").Sort
        .SetRange alan 'Makro recordardan sabit gelen Range("A2:C6")'i değiştirdik
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
'----------------------------------------------------

Sub resizeiledinamiksıralama2()
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
End Sub
'----------------------------------------------------

Sub resize_range_diziata()
Dim sayılar As Variant
sayılar = [A1].CurrentRegion.Value
'colon sayısı:ubound(siciller,2)
'row sayısı:Ubound(siciller,1) veya kısaca ubound(siciller), 1 defaulat value
Range("A8").Resize(UBound(sayılar), UBound(sayılar, 2)).Value = sayılar
End Sub
'----------------------------------------------------

Sub ZebraYa_rows()
Dim alan As Range

Set alan = Range("A2:C6")
For i = 1 To alan.Rows.Count
    If i Mod 2 = 1 Then
        alan.Rows(i).Interior.Color = vbYellow
    Else
        alan.Rows(i).Interior.Color = vbWhite
    End If
Next i

With alan.Columns(1)
    .Font.Bold = True
    .Font.Color = vbRed
End With 
End Sub
'----------------------------------------------------

Sub silme_ekleme_hide()
'macro recorder
End Sub

Sub findreplace()
'recorder
'lookat dikkat
Dim arabul As Range
Set arabul = Range("A1").CurrentRegion.Find(102)
arabul.Select

Set arabul = Range("A1").CurrentRegion.Find( _
        What:="volkan", _
        After:=ActiveCell, _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
End Sub
'----------------------------------------------------

Sub tumworkbooktaFIND()
'Tüm sayfalarda arama
Dim ws As Worksheet
Dim arabul As Range

For Each ws In ActiveWorkbook.Sheets
    ws.Activate
    Set arabul = Cells.Find( _
        What:="volkan", _
        After:=ActiveCell, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    If Not arabul Is Nothing Then
        arabul.Select
        Exit For
    End If
Next ws
End Sub
'----------------------------------------------------

Sub replaceornek()
Cells.Replace _
        What:="ali", _
        Replacement:="veli", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
'----------------------------------------------------

Sub rangecalc()
    Selection.Calculate 'Sadece seçili hücrelerdeki formülleri hesaplatır.
End Sub
'----------------------------------------------------

Sub parentornek()

Debug.Print TypeName(ActiveCell.Parent) 'Worksheet
Debug.Print ActiveCell.Parent.Name 'ilgili Worksheetin adı
End Sub
