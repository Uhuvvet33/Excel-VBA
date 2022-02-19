'-----------------Farklı excel dosyalarını sayfa olarak birleştirmek. (Bu alıntıdır)

Sub SayfaEkle()
' Sayfaları Ekleme
Application.ScreenUpdating = False  ' Ekran titremesini iptal

Dim Filename As String

' Aşağıdaki Dosya Yolu olarak Masaüstü ve Sayfa1 K5 hücresinden alınıyor
Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator & _
        Sheets(1).Range("k5").Value & "\"
' Dosya Yolunu Sayfa bir K5 hücresinden alıyor
Filename = Dir(Path & "*.xlsx")
say = 1

Do While Filename <> ""
    Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
    For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=ThisWorkbook.Sheets(1)
        ' Sayfa adı ile dosya adı aynı olması için yazdığımız kod
        Dim LArray() As String
        LArray = Split(Filename, ".")
        Sheets(2).Name = Mid(LArray(0), 1, 21) & " (" & say & ")"
    Next Sheet
    Workbooks(Filename).Close
    Filename = Dir()
    say = say + 1
    
Loop

Sheets(1).Range("K6").Value = Sheets.Count - 1  ' Eklenen Sayfa Sayısı
Sheets(1).Range("K7").Value = Format(Now)       ' İşlem Zamanı
Sheets(1).Select

End Sub
