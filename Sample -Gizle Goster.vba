
Sub EkranTitreme()
  'Komutlar çalıştırlırken ekranda beliren titremeleri kapatır.
  Application.ScreenUpdating = False  
End Sub  
'----------------------------------------------------------

Sub GizleGoster()           'Satır gizle, göster
'Gizle; Belirlenen sayfa ve hücrelerdeki boş hücreleri
  For Each t In Worksheets("Sayfa1").Range("c6:c38").Cells
    If t.Value = "" Then          'Hücre boş ise
    t.EntireRow.Hidden = True     'Gizle
  End If
  Next t

'Göster; Belirlenen sayfa ve hücrelerdeki boş hücreleri
  For Each t In Worksheets("Kursiyer").Range("c6:c38").Cells
    If t.Value = "" Then        
    t.EntireRow.Hidden = False    'Göster
  End If
  Next t
End Sub
