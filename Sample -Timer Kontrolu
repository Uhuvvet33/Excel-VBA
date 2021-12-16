Sub TimerKontrolu()
' Değişken tanımlamanın verimli kullanım örneği,
' İşin ne kadar sürdüğünü hesaplama
Dim basla As Single
Dim bitis As Single
Dim i As Long

basla = Timer               'Saatin değerini alıyoruz
For i = 1 To 100000000
    k = k + 1
Next i
bitis = Timer               'Döngü bittikten sonra bir kez daha zaman değerini alıyoruz

MsgBox ("İşlem Süresi :" & vbNewLine & Round(bitis - basla, 2) & " saniyedir.")
End Sub
