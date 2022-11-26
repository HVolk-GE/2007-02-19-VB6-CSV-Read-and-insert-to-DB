Attribute VB_Name = "MakleQueryCSVToDB"
Public Function SearchFileForText(ByVal sFile As String, _
ByVal sText As String, _
Optional ByVal lngStart As Long = 1) As Long

Dim F As Integer
Dim lngStrLen As Long
Dim lngFound As Long
Dim lngFileSize As Long
Dim lngFilePos As Long
Dim lngReadSize As Long
Dim sTemp As String
Dim sPrev As String
Dim intProz As Integer

' Größe eines einzelnen einzulesenden Datenblocks
Const lngBlockSize = 4096

' Länge des gesuchten Textes
lngStrLen = Len(sText)

' Falls die Datei gar nicht existiert, oder der
' kein Suchtext angegeben wurde, wird die Funktion
' hier verlassen
If Dir$(sFile) = "" Or lngStrLen = 0 Then Exit Function

' Datei im Binärmodus öffnen
F = FreeFile

Open sFile For Binary As #F

' Größe der Datei
lngFileSize = LOF(F)

' Start-Position
If lngStart > 1 Then
    Seek #F, lngStart
    lngFilePos = lngStart - 1
End If

' Solange "blockweise" einlesen, bis entweder das
' Dateiende erreicht oder der Text gefunden wurde
While lngFilePos < lngFileSize And lngFound = 0
    
    If lngFilePos + lngBlockSize > lngFileSize Then
    ' Falls aktuelle Position + Blockgröße über das
    ' Dateiende hinaus geht -> Blockgröße neu festlegen
    ' (maximal bis Dateiende)
        lngReadSize = lngFileSize - lngFilePos
    Else
    ' ansonsten: festgelegte Blockgröße einlesen
        lngReadSize = lngBlockSize
    End If
    ' Variable vorbereiten (mit Leerzeichen fülen)
    sTemp = Space$(lngReadSize)
    
    ' Datenblock einlesen (Größe = lngReadSize)
    Get #F, , sTemp
    
    ' die letzten Zeichen des vorigen Blocks nochmals
    ' mit in den Suchvorgang einbeziehen, denn es
    ' könnte ja sein, dass sich der gesuchte Text
    ' genau an zwischen dem letzten und dem aktuell
    ' eingelesenen Block befindet
    sTemp = sPrev + sTemp
    
    ' Ist der gesuchte Text enthalten?
    lngFound = InStr(sTemp, sText)
    
    If lngFound > 0 Then
        ' JA, Suchtext ist enthalten!
        ' Position ermitteln
        lngFound = lngFilePos + lngFound - lngStrLen
    End If
    
    ' aktuelle Position aktualisieren
    lngFilePos = lngFilePos + lngReadSize
    ' Fortschritt anzeigen
    'intProz = Int(lngFilePos / lngFileSize * 100 + 0.5)
    'lblStatus.Caption = "Suche läuft... " & CStr(intProz) & "%"

    DoEvents
    sPrev = Right$(sTemp, lngStrLen)
Wend
    ' nachfolgender Code nur zu Testzwecken
    ' (einfach später dann auskommentieren)
    If lngFound > 0 Then
        sTemp = Space$(lngStrLen)
        Seek #F, lngFound
        Get #F, , sTemp
        Debug.Print sTemp
    End If
    
' Datei schliessen
Close #F

' Funktionsrückgabewert: Fundstelle (Position)
' SearchFileForText = lngFound

End Function

Sub MakeDefaultCSVToDB()
    rowcnt0 = 0
    Form1.Command1.Enabled = False
    Form1.Command2.Enabled = False
    
    
    Form1.File1.Path = Form1.Text1.Text
    Form1.File1.FileName = Form1.Combo1.Text
    
    Form1.File1.Refresh
    
    SourcePath = Form1.Text1.Text
    StrDesFile2 = Form1.Text2.Text

    For g = 0 To Form1.File1.ListCount - 1
        
        Form1.ProgressBar1.Visible = True
        
        intProgress = Form1.File1.ListCount + 1
    
        intValueFiles = 100 / intProgress

        StrDesFile2 = SourcePath & Form1.File1.List(g)
  
        Open StrDesFile2 For Input As 1
     
        Line Input #1, FileNames(3, 0)

        FileNames(3, 0) = Replace(FileNames(3, 0), ";", ", ")

'###############################################################################
        
        sTable = ""
        
        Form1.ProgressBar1.Value = intValueFiles * 1
        
        While Not EOF(1)
            Line Input #1, FileNames1(rowcnt0)
            If sTable = "" Then
                For l = 1 To Len(FileNames1(rowcnt0))
                    strTmp0 = Mid(FileNames1(rowcnt0), l, 1)
                    If strTmp0 = ";" Then
             
                    temptxt00 = UCase(Trim(Mid(FileNames1(rowcnt0), l + 1, 2)))
             
                        If temptxt00 = "A-" Or temptxt00 = "C-" Or temptxt00 = "H-" Or _
                           temptxt00 = "P-" Or temptxt00 = "S-" Or temptxt00 = "V-" Or _
                           temptxt00 = "W-" Then
                           
                           If temptxt00 = "A-" Then
                              sTable = "alpine"
                           ElseIf temptxt00 = "C-" Then
                              sTable = "crack"
                           ElseIf temptxt00 = "H-" Then
                              sTable = "homologation"
                           ElseIf temptxt00 = "P-" Then
                              sTable = "performance"
                           ElseIf temptxt00 = "S-" Then
                              sTable = "strength"
                           ElseIf temptxt00 = "V-" Then
                              sTable = "various"
                           ElseIf temptxt00 = "W-" Then
                              sTable = "wear"
                           End If
                           
'########################################################################################
                            ' Table Name changed :
                            VERSUCH = LCase(Trim(Mid(FileNames1(rowcnt0), l + 1, 8)))
                            'VERSUCH = Replace(VERSUCH, "-", "_")
                            If InStr(1, VERSUCH, "-") Or InStr(1, VERSUCH, "_") _
                               Then
                               sTable = "`" & VERSUCH & "`"
                            Else
                               sTable = VERSUCH
                            End If
                            If TBLCreate = 0 Then
                               CheckTablesNames
                            End If
'########################################################################################
                           dbConnect
                           a = 0
                        Exit For
                        End If
                    End If
                Next l
            End If
            
            If DBUSDEF = "YES" Then
               FileNames1(rowcnt0) = Replace(FileNames1(rowcnt0), ",", ".")
            End If
            
            FileNames1(rowcnt0) = "('" & Replace(FileNames1(rowcnt0), ";", "', '") & "')"
 
 '###############################################################################
           
            rowcnt0 = rowcnt0 + 1
           
            intVal = g + 1
            
            If rowcnt0 = 98 Then
               Form1.ProgressBar1.Value = intValueFiles * intVal / 2
               DoEvents
               dbRowAdd
               rowcnt0 = 0
            End If
        Wend
        
        If rowcnt0 <> 0 Then
           dbRowAdd
           rowcnt0 = 0
        End If
        
        Close #1
        
    Next g
     
    dbDisconnect
     
     For h = 0 To Form1.File1.ListCount - 1
        
        Kill SourcePath & Form1.File1.List(h)
     
     Next h
     
    Form1.Combo2.Refresh
    Form1.Combo2.Text = ""
    Form1.ProgressBar1.Visible = False
    Form1.Command1.Enabled = False
    Form1.Command2.Enabled = True
    'If MESAG = "YES" Then
       MsgBox "Export to database done...!"
    'End If
    End
End Sub
