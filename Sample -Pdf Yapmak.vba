
Sub Pdf_Yap
DosyaYolu = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator   'Dosya yolu olarak Masaüstü belirleniyor.
'DosyaYolu = Worksheets("BilgiGirisi").Range("M2").Text   'Dosyayolunun hücre üzerinden alınması

BasTarihi = Worksheets("BilgiGirisi").Range("c5").Value   'Dosya isminin belirlenmesinde hücre üzerinde yer alan tarih bilgisi alınıyor.

'Farklı sayfalardaki listeler pdf olarak yazdırılıyor
Worksheets("Sayfa1").ExportAsFixedFormat xlTypePDF, Filename:=DosyaYolu & BasTarihi & " Liste.pdf"
Worksheets("Sayfa2").ExportAsFixedFormat xlTypePDF, Filename:=DosyaYolu & BasTarihi & " Liste.pdf"
End Sub
