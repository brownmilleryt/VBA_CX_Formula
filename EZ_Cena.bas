Function EZ_Cena(receptura As String, Optional kod As String = "K069", Optional FuncExit As Boolean = False) As Variant

Dim objIE As Object
Dim błąd As Variant
Dim cena As Variant

receptura = Replace(receptura, "/", "%2F")

the_start:

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Top = 0
objIE.Left = 0
objIE.Width = 800
objIE.Height = 600
objIE.Visible = True

On Error Resume Next
objIE.navigate ("http:##########/mt_receptura?receptura=" & receptura & "&wuid=" & zakład(kod))

Do
DoEvents
       If Err.Number <> 0 Then
               objIE.Quit
               Set objIE = Nothing
               GoTo the_start:
       End If
Loop Until objIE.readystate = 4

cena = objIE.document.getElementById("kalkulacja_przeliczana_wynik").InnerText
błąd = objIE.document.getElementsByClassName("crit")(0).InnerText

If Not IsEmpty(cena) Then
    EZ_Cena = Round(CSng(Replace(Trim(cena), ".", ",")), 2)
Else
    If Not IsEmpty(błąd) Then
        EZ_Cena = "Receptura nie istnieje lub jest nieaktualna"
        objIE.Quit
    Else
        EZ_Cena = "Zaloguj się do systemu EZ i uruchom ponownie formułę"
        objIE.Visible = True
    End If
End If

If FuncExit Then
    objIE.Quit
End If

End Function
