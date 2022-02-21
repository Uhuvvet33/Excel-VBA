' TC Kontrolü Fonksiyon olarak.
' Örnek Alıntıdır.

Function TCKontrol(ByVal tc As String) As Boolean

Dim tc1 As Integer
Dim tc2 As Integer
Dim tc3 As Integer
Dim tc4 As Integer
Dim tc5 As Integer
Dim tc6 As Integer
Dim tc7 As Integer
Dim tc8 As Integer
Dim tc9 As Integer
Dim tc10 As Integer
Dim tc11 As Integer
Dim tc10x As Integer
Dim tc11x As Integer

tc1 = Mid(tc, 1, 1)
tc2 = Mid(tc, 2, 1)
tc3 = Mid(tc, 3, 1)
tc4 = Mid(tc, 4, 1)
tc5 = Mid(tc, 5, 1)
tc6 = Mid(tc, 6, 1)
tc7 = Mid(tc, 7, 1)
tc8 = Mid(tc, 8, 1)
tc9 = Mid(tc, 9, 1)
tc10 = Mid(tc, 10, 1)
tc11 = Mid(tc, 11, 1)

tc10x = ((tc1 + tc3 + tc5 + tc7 + tc9) * 7 - (tc2 + tc4 + tc6 + tc8)) Mod 10
tc11x = (tc1 + tc2 + tc3 + tc4 + tc5 + tc6 + tc7 + tc8 + tc9 + tc10x) Mod 10

If tc10 = tc10x And tc11 = tc11x Then
    TCKontrol = True
Else
    TCKontrol = False
End If

End Function
