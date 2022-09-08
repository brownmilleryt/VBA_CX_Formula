Function EZ_Sap(receptura As String, Optional kod As String = "K069", Optional FuncExit As Boolean = False) As String


Dim objIE As Object
Dim Sap As Variant


receptura = Replace(receptura, "/", "%2F")

the_start:

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Top = 0
objIE.Left = 0
objIE.Width = 800
objIE.Height = 600
objIE.Visible = True

On Error Resume Next
objIE.navigate ("#################mt_receptura?receptura=" & receptura & "&wuid=" & zakład(kod))

Do
DoEvents
       If Err.Number <> 0 Then
               objIE.Quit
               Set objIE = Nothing
               GoTo the_start:
       End If
Loop Until objIE.readystate = 4

Sap = objIE.document.getElementsByClassName("hint")(0).InnerText

Sap = Split(Sap, "-")

If Len(Sap(0)) > 1 Then
    EZ_Sap = Trim(Sap(0))
Else
    EZ_Sap = "Brak numeru SAP"
    objIE.Quit
End If

If EZ_Sap = "" Then
   EZ_Sap = "Brak receptury"
   objIE.Quit
End If

If FuncExit Then
    objIE.Quit
End If



End Function
