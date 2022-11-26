Attribute VB_Name = "Module2"
Public Function HexToStr(sHexString As String) As String
Dim ai As Integer, iHigh, Value, Value2 As Integer, iLow As Integer
Dim sHexChar, Value3 As String
Dim sLow As String, sHigh As String

If LenB(" ") = Len(" ") Then ' Unicode-System?
   HexToStr = Space(Len(sHexString) / 3)
Else
   HexToStr = Space(Len(sHexString) / 6)
End If

Value = (Len(sHexString) * 256) + 1
'Value2 = (Value -

If Value > 32768 Then Value = Value2 - 65536

Value2 = Asc(Left(sHexString, Value)) ' Value))
Value3 = Asc(Mid(sHexString, Value2 + 1, 2048))

For ai = 1 To Len(sHexString) Step 3 ' aus 3 Zeichen wird 1 Byte!
    iHigh = Asc(Mid(sHexString, ai, 1)) - 48 ' Annahme: Ziffer
    If iHigh > 15 Then iHigh = iHigh - 7    ' War es etwa ein Buchstabe?
    iLow = Asc(Mid(sHexString, ai + 1, 1)) - 48
    
    If iLow > 15 Then iLow = iLow - 7
    
    MidB(HexToStr, (i + 2) \ 3) = ChrB(iHigh * 16 + iLow) ' Zusammensetzen
    
Next ai

End Function
