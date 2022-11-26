Attribute VB_Name = "Module4"
Sub MyTest01ReadMicrofaceResults(ResFile)
Dim AVGValues(), INames(), IScale(), IOffset() 'AVGValues, INames, Scale , Offset
Dim VERSUCH, TDNr1, TDNr2, rRadius, Inertia As String
Dim frictionconfact1, frictionconfakt2, CurrentBlock As String
Dim CanelCnt As Integer

    F = FreeFile
    
    FileNames(4, 0) = ""
    
    Open ResFile For Input As #F

    While Not EOF(F)
        Line Input #F, c1
        c1 = Replace(c1, " ", "")
        c1 = Trim(c1)
        If Left(c1, 2) = "0," Then
          If TNr = "" Then
            Line Input #F, temp ' Results File Name
            TNr = temp
            Line Input #F, temp ' Date
            Line Input #F, temp ' Schedule
            VERSUCH = temp      ' Write Schedule Nr to Variable for use this in next time
            temptxt00 = Left(temp, 2) ' Write Schedule Nr to Categorie Variable first 2nd Charaters
            DBTableNames
            Line Input #F, temp ' Project End 1
            TDNr = temp         ' Write Project Data to Variables1
            Line Input #F, temp ' Project End 2
            TDNr2 = temp        ' Write Project Data to Variables2 not use on Bad Camberg
            Line Input #F, temp ' Rolling Radius mm
            rRadius = temp       ' Write Rolling Radius to Variable
            Line Input #F, temp ' Inertia Kgm2
            Inertia = temp      ' Write Inertia to Variable
            Line Input #F, temp ' Friction Convesion Factor End 1
            frictionconfakt1 = temp ' Write Friction Convesion End 1 to Variables1
            Line Input #F, temp ' Friction Convesion Factor End 2
            frictionconfakt2 = temp ' Write Friction Convesion End 2 to Variables2
        
            'StrSoucreFile4 = StrDesFile2 & ACdataFile
            'StrSoucreFile4 = StrSoucreFile4 & "_" & TNr & ".csv"
            'Open StrSoucreFile4 For Output As 6
           End If
        ElseIf Left(c1, 2) = "1," Then
            Input #F, bl    ' Logged Block Lenght
            CurrentBlock = String$(bl, " ")
            fp = Seek(F)
            Close #F
            Open ResFile For Binary As F
            Get #F, fp, CurrentBlock    ' Logged Block
            fp = Seek(F)
            Close #F
            Open ResFile For Input As F
            Seek #F, fp
            Line Input #F, temp   ' past the cr/lf
            Line Input #F, temp   ' Logging rate
            i = 0
            For l = 1 To bl - 2 Step 3
                cn = Asc(Mid$(CurrentBlock, l, 1))  ' Channel Number
                If cn <> 99 Then    ' Normal Channel
                    Value = Asc(Mid$(CurrentBlock, l + 1, 1)) + (Asc(Mid$(CurrentBlock, l + 2, 1)) * 256#) ' Value
                    If Value > 32767# Then
                        Value = Value - 65536#
                    End If
                    Value = (Value - IOffset(cn)) * IScale(cn)  ' Value In Real units
                Else
                    cn = Asc(Mid$(CurrentBlock, l + 1, 1))
                    If l < Len(CurrentBlock) - 4 Then
                        Value = Asc(Mid$(CurrentBlock, l + 2, 1)) + (Asc(Mid$(CurrentBlock, l + 3, 1)) * 256#) + (Asc(Mid$(CurrentBlock, l + 4, 1)) * 65535#) + (Asc(Mid$(CurrentBlock, l + 5, 1)) * 16777216)
                        l = l + 3
                    End If
                End If
                
            Next l
        ElseIf Left(c1, 2) = "2," Then
               
               For i = 1 To Len(c1)
                    tmpchr = ","
                    tmpchar = Mid(c1, i, 1)
                    If tmpchar = tmpchr Then
                        c2 = Mid(c1, i + 1, Len(c1) - i + 1)
                        Exit For
                    End If
               Next i
               
               CanelCnt = Trim(c2)
               CanelCnt = CInt(CanelCnt)
               c2 = CInt(c2)
               ReDim Preserve INames(CanelCnt, 0)
            For l = 1 To c2
                ReDim Preserve IScale(l)
                ReDim Preserve IOffset(l)
                Line Input #F, INames(l, 0) ' Instrument Name Use by Columnsnames
                Line Input #F, IScale(l) ' Scale
                Line Input #F, IOffset(l) ' Offset
                IScale(l) = Replace(IScale(l), ".", ",")
                IScale(l) = CDec(IScale(l))
                IOffset(l) = Replace(IOffset(l), ".", ",")
                IOffset(l) = CDec(IOffset(l))
            Next l
            
            For l = 1 To c2 - 1
                WriteZeile = WriteZeile & INames(l, 0) & ";"
            Next l
                WriteZeile = WriteZeile & INames(c2, 0)
            
            'Print #6, WriteZeile & ";" & "SEQUENCE;PRUEFLING;SCHLUESSEL;VERSUCH"
            
        ElseIf Left(c1, 2) = "3," Then
            Line Input #F, temp ' Sequence Name
            Line Input #F, temp ' Sequence Number
            Line Input #F, temp '
        ElseIf Left(c1, 2) = "4," Then
            Line Input #F, d1    ' Total Stop Time
            Line Input #F, d2    ' Squence Time
        ElseIf Left(c1, 2) = "5," Then
'#
            Line Input #F, temp ' Test Referance
            Line Input #F, temp ' Material
        ElseIf Left(c1, 2) = "6," Then
        'Case 6  ' Average Data
            For va_loop = 1 To 16
                Line Input #F, temp  ' Average Data
            Next va_loop
        ElseIf Left(c1, 2) = "7," Then
            Line Input #F, temp  ' Brake Time Secs
            Line Input #F, temp  ' Revs to Stop
        ElseIf Left(c1, 2) = "8," Then
        'Case 8  ' Results Format
            Line Input #F, temp
        ElseIf Left(c1, 3) = "10," Then
        'Case 10 ' Quality Values Info
            For l = 1 To 40
                Line Input #F, temp ' Quality Value Info
            Next l
        ElseIf Left(c1, 3) = "11," Then
        'Case 11 ' Quality String Info
            For l = 1 To 6
                Line Input #F, temp ' Quality String Info
            Next l
        ElseIf Left(c1, 3) = "12," Then
        'Case 12 ' ATE Info
            Line Input #F, temp  ' Servo Channel Number
            Line Input #F, temp  ' Servo Demand Value
        ElseIf Left(c1, 3) = "13," Then
        'Case 13 ' MFDD + Extras
            Line Input #F, d1    ' MFDD
            For l = 1 To 5
                Line Input #F, d1
            Next l
        ElseIf Left(c1, 3) = "20," Then
        'Case 20 ' Torque Offset ****(Not Used)****
            '#Line Input #F, temp
            '#Input #F, d1
            '#If c2 = 1 Then ' end 1
            '#    IOffset(Torque1) = d1
            '#Else ' end 2
            '#    IOffset(Torque2) = d1
            '#End If
        ElseIf Left(c1, 3) = "21," Then
        'Case 21 ' thermocouple bits for Roulunds  ****(Not Used)****
            For l = 1 To 18
                Line Input #F, d1
            Next l
        ElseIf Left(c1, 3) = "22," Then
        'Case 22 ' Running ISim
            Line Input #F, d1    ' Inertia Simulation 1 = On 0 = Off
        ElseIf Left(c1, 3) = "23," Then
        'Case 23 ' Comments
            Line Input #F, temp ' Comments
        ElseIf Left(c1, 3) = "24," Then
        'Case 24 ' Measurements
            For l = 1 To 17
                Line Input #F, d1    ' Measurements
            Next l
        ElseIf Left(c1, 3) = "25," Then
        'Case 25 ' Measurements
            For l = 1 To 26
                Line Input #F, d1    ' Measurements
            Next l
        ElseIf Left(c1, 3) = "26," Then
        'Case 26 ' Inertia
            Line Input #F, d1
        ElseIf Left(c1, 3) = "27," Then
        'Case 27 ' Run Test Channel Information  ****(Not Used)****
            Line Input #F, d1 ', d1, d1, d1
        ElseIf Left(c1, 3) = "28," Then
        'Case 28 ' Time and Date Info
            Line Input #F, d1
            Line Input #F, d1
            Line Input #F, d1
        ElseIf Left(c1, 3) = "29," Then
        'Case 29 ' 6 Lines of Comments
            For l = 1 To 6
                Line Input #F, temp ' Comments
            Next l
        ElseIf Left(c1, 3) = "30," Then
        'Case 30 ' sequence start date and time
            Line Input #F, temp ' number
            Line Input #F, temp ' name
            Line Input #F, temp ' date
            Line Input #F, temp ' time
       End If

    Wend
End Sub

Sub ReadMicrofaceResultsOrg(ResFile)
Dim AVGValues(), INames(), IScale(), IOffset() 'AVGValues, INames, Scale , Offset
Dim VERSUCH, TDNr1, TDNr2, rRadius, Inertia As String
Dim frictionconfact1, frictionconfakt2, CurrentBlock As String
Dim CanelCnt As Integer

    F = FreeFile
    
    FileNames(4, 0) = ""
    
    Open ResFile For Input As #F
    
    While Not EOF(F)
        Input #F, c1, c2
        Select Case c1
        Case 0
          If TNr = "" Then
            Line Input #F, temp ' Results File Name
            TNr = temp
            Line Input #F, temp ' Date
            Line Input #F, temp ' Schedule
            VERSUCH = temp      ' Write Schedule Nr to Variable for use this in next time
            temptxt00 = Left(temp, 2) ' Write Schedule Nr to Categorie Variable first 2nd Charaters
            DBTableNames
            Line Input #F, temp ' Project End 1
            TDNr = temp         ' Write Project Data to Variables1
            Line Input #F, temp ' Project End 2
            TDNr2 = temp        ' Write Project Data to Variables2 not use on Bad Camberg
            Line Input #F, temp ' Rolling Radius mm
            rRadius = temp       ' Write Rolling Radius to Variable
            Line Input #F, temp ' Inertia Kgm2
            Inertia = temp      ' Write Inertia to Variable
            Line Input #F, temp ' Friction Convesion Factor End 1
            frictionconfakt1 = temp ' Write Friction Convesion End 1 to Variables1
            Line Input #F, temp ' Friction Convesion Factor End 2
            frictionconfakt2 = temp ' Write Friction Convesion End 2 to Variables2
        
            StrSoucreFile4 = StrDesFile2 & ACdataFile
            StrSoucreFile4 = StrSoucreFile4 & "_" & TNr & ".csv"
            Open StrSoucreFile4 For Output As 6
           End If
        Case 1
            Input #F, bl    ' Logged Block Lenght
            CurrentBlock = String$(bl, " ")
            fp = Seek(F)
            Close #F
            Open ResFile For Binary As F
            Get #F, fp, CurrentBlock    ' Logged Block
            fp = Seek(F)
            Close #F
            Open ResFile For Input As F
            Seek #F, fp
            Line Input #F, temp   ' past the cr/lf
            Line Input #F, temp   ' Logging rate
           ' ReDim Preserve INames(i, CInt(temp))
            i = 0
            For l = 1 To bl - 2 Step 3
                cn = Asc(Mid$(CurrentBlock, l, 1))  ' Channel Number
                If cn <> 99 Then    ' Normal Channel
                    Value = Asc(Mid$(CurrentBlock, l + 1, 1)) + (Asc(Mid$(CurrentBlock, l + 2, 1)) * 256#) ' Value
                    If Value > 32767# Then
                        Value = Value - 65536#
                    End If
                    Value = (Value - IOffset(cn)) * IScale(cn)  ' Value In Real units
                Else
                    cn = Asc(Mid$(CurrentBlock, l + 1, 1))
                    If l < Len(CurrentBlock) - 4 Then
                        Value = Asc(Mid$(CurrentBlock, l + 2, 1)) + (Asc(Mid$(CurrentBlock, l + 3, 1)) * 256#) + (Asc(Mid$(CurrentBlock, l + 4, 1)) * 65535#) + (Asc(Mid$(CurrentBlock, l + 5, 1)) * 16777216)
                        l = l + 3
                    End If
                End If
                
            Next l
         Case 2
               CanelCnt = c2
               ReDim Preserve INames(CanelCnt, 0)
            For l = 1 To c2
                 ReDim Preserve IScale(l)
                ReDim Preserve IOffset(l)
                Line Input #F, INames(l, 0) ' Instrument Name Use by Columnsnames
                Line Input #F, IScale(l) ' Scale
                Line Input #F, IOffset(l) ' Offset
                IScale(l) = Replace(IScale(l), ".", ",")
                IScale(l) = CDec(IScale(l))
                IOffset(l) = Replace(IOffset(l), ".", ",")
                IOffset(l) = CDec(IOffset(l))
            Next l
            
            For l = 1 To c2 - 1
                WriteZeile = WriteZeile & INames(l, 0) & ";"
            Next l
                WriteZeile = WriteZeile & INames(c2, 0)
            
            'Print #6, WriteZeile & ";" & "AC" & strDemilier & _
                      "TDNr" & strDemilier & TNr & strDemilier & VERSUCH
             Print #6, WriteZeile & ";" & "SEQUENCE;PRUEFLING;SCHLUESSEL;VERSUCH"
            
        Case 3
            Line Input #F, temp ' Sequence Name
       '     FileNames(4, 0) = temp
            Line Input #F, temp ' Sequence Number
     '       FileNames(4, 0) = temp
            Line Input #F, temp '
        Case 4
            Input #F, d1    ' Total Stop Time
            Input #F, d2    ' Squence Time
        Case 5
            Line Input #F, temp ' Test Referance
            Line Input #F, temp ' Material
        Case 6  ' Average Data
            If imt = 0 Then imt = 1
            'Debug.Print FileNames(4, 0)
            For va_loop = 1 To 16
                Input #F, temp  ' Average Data
               ' If FileNames(4, imt) <> "" Then
               '    FileNames(4, imt) = FileNames(4, imt) & ";" & temp
               ' Else
               '    FileNames(4, imt) = temp
               ' End If
            Next va_loop
            'Debug.Print FileNames(4, imt)
            imt = imt + 1
        Case 7
            Input #F, temp  ' Brake Time Secs
            Input #F, temp  ' Revs to Stop
        Case 8  ' Results Format
            Line Input #F, temp
        Case 10 ' Quality Values Info
            For l = 1 To 40
                Line Input #F, temp ' Quality Value Info
            Next l
        Case 11 ' Quality String Info
            For l = 1 To 6
                Line Input #F, temp ' Quality String Info
            Next l
        Case 12 ' ATE Info
            Input #F, temp  ' Servo Channel Number
            Input #F, temp  ' Servo Demand Value
        Case 13 ' MFDD + Extras
            Input #F, d1    ' MFDD
            For l = 1 To 5
                Input #F, d1
            Next l
        Case 20 ' Torque Offset ****(Not Used)****
            Line Input #F, temp
            Input #F, d1
            If c2 = 1 Then ' end 1
                IOffset(Torque1) = d1
            Else ' end 2
                IOffset(Torque2) = d1
            End If
        Case 21 ' thermocouple bits for Roulunds  ****(Not Used)****
            For l = 1 To 18
                Input #F, d1
            Next l
        Case 22 ' Running ISim
            Input #F, d1    ' Inertia Simulation 1 = On 0 = Off
        Case 23 ' Comments
            Line Input #F, temp ' Comments
        Case 24 ' Measurements
            For l = 1 To 17
                Input #F, d1    ' Measurements
            Next l
        Case 25 ' Measurements
            For l = 1 To 26
                Input #F, d1    ' Measurements
            Next l
        Case 26 ' Inertia
            Input #F, d1
        Case 27 ' Run Test Channel Information  ****(Not Used)****
            Input #F, d1, d1, d1, d1
        Case 28 ' Time and Date Info
            Input #F, d1
            Input #F, d1
            Input #F, d1
        Case 29 ' 6 Lines of Comments
            For l = 1 To 6
                Line Input #F, temp ' Comments
            Next l
        Case 30 ' sequence start date and time
            Line Input #F, temp ' number
            Line Input #F, temp ' name
            Line Input #F, temp ' date
            Line Input #F, temp ' time
        End Select
    Wend
End Sub

'###########################################################################
Sub ReadMicrofaceResults(ResFile)
Dim IOffset(), IScale(), Intrument() As String
Dim CurrentBlock As String
Dim oldseqNr As String
Dim cntSeqName As Long
Dim CntIntr As Integer

cntSeqName = 5
cntrow = 1
rowcnt = 1
ReDim Preserve Intrument(0)
CntIntr = 0

Intrument(0) = ""

    F = FreeFile
    
    Open ResFile For Input As #F
    
    While Not EOF(F)
        Input #F, c1, c2
        Select Case c1
        Case 0

            'Test Date
            temp = "Test Number"
            temp = "Project Data Left"
            temp = "Project Data Right"
            temp = "Pad Description Left"
            temp = "Pad Description Right"
            temp = "Conversion Format"
            temp = "Test Date"
            Line Input #F, temp ' Results File Name
            temp = temp
            Line Input #F, temp ' Date
            temp = temp
            Line Input #F, temp ' Schedule
            'Worksheets(1).Range("C1").Value = temp
            Line Input #F, temp ' Project End 1
            temp = temp
            Line Input #F, temp ' Project End 2
            temp = temp
            Line Input #F, temp ' Rolling Radius mm
            'Worksheets(1).Range("F1").Value = CDbl(temp)
            Line Input #F, temp ' Inertia Kgm2
            'Worksheets(1).Range("G1").Value = temp
            Line Input #F, temp ' Friction Convesion Factor End 1
            'FrictionConvFac1 = temp
            Line Input #F, temp ' Friction Convesion Factor End 2
            'FrictionConvFac1 = temp
            'End If
        Case 1
            Input #F, bl    ' Logged Block Lenght
            'Debug.Print "Inhalt bl: " & bl ' Hier tritt bei dem RES File immer in Sequence 4 ein Fehler auf
                           ' Typ unzulässig !
            'Debug.Print "Länge bl: " & Len(bl)
            
            If Len(bl) > 5 Then
               bl = Left(bl, 4)
               bl = Trim(bl)
            End If
            
            CurrentBlock = String$(bl, " ")
            fp = Seek(F)
            Close #F
            Open ResFile For Binary As F
            Get #F, fp, CurrentBlock    ' Logged Block
            fp = Seek(F)
            Close #F
            Open ResFile For Input As F
            Seek #F, fp
            Line Input #F, temp   ' past the cr/lf
            Line Input #F, temp   ' Logging rate
            For l = 1 To bl - 2 Step 3
                cn = Asc(Mid$(CurrentBlock, l, 1))  ' Channel Number
                
                If cn > 0 Then
                
                    If cn = 1 Then
                       cntrow = cntrow + 1
                    End If
                  
                    If cn <= CntIntr Then
                        If cn <> 99 Then    ' Normal Channel
                            Value = Asc(Mid$(CurrentBlock, l + 1, 1)) + (Asc(Mid$(CurrentBlock, l + 2, 1)) * 256#) ' Value
                            If Value > 32767# Then
                                Value = Value - 65536#
                            End If
                            Value = ((Value - IOffset(cn)) * IScale(cn))  ' Value In Real units
                        Else
                            cn = Asc(Mid$(CurrentBlock, l + 1, 1))
                            If l < Len(CurrentBlock) - 4 Then
                                Value = Asc(Mid$(CurrentBlock, l + 2, 1)) + (Asc(Mid$(CurrentBlock, l + 3, 1)) * 256#) + (Asc(Mid$(CurrentBlock, l + 4, 1)) * 65535#) + (Asc(Mid$(CurrentBlock, l + 5, 1)) * 16777216)
                                l = l + 3
                            End If
                        End If

                        Value = Replace(Value, ".", ",")
                        'Cells(cntrow, cn + 2).Select
                        'Selection.NumberFormat = "0.000"
                        'Cells(cntrow, cn + 2).Value = CDbl(Value)
                        'Cells(cntrow, 1).Value = SequenceNr
                    
                        If cntrow > 65535 Then
                            temp = SequenceNr & "a"
                            If cn = 1 Then
                                cntrow = 1
                            Else
                                cntrow = 2
                            End If
                        End If
                    End If
                    rowcnt = cn
                End If
            Next l
        Case 2
           If Intrument(0) = "" And CntIntr = 0 Then
                CntIntr = c2
                For l = 1 To c2
                    ReDim Preserve Intrument(l)
                    ReDim Preserve IScale(l)
                    ReDim Preserve IOffset(l)
                    Line Input #F, Intrument(l) ' temp ' Instrument Name
                    Line Input #F, IScale(l) ' Scale
                    Line Input #F, IOffset(l) ' Offset
                    IScale(l) = Replace(IScale(l), ".", ",")
                    IScale(l) = CDec(IScale(l))
                    'Debug.Print IScale(l)
                    IOffset(l) = Replace(IOffset(l), ".", ",")
                    IOffset(l) = CDec(IOffset(l))
                Next l
            End If
        Case 3
            Line Input #F, temp ' Sequence Name
            SequenceName = temp
            Line Input #F, temp ' Sequence Number
            SequenceNr = Trim(temp)
            
            If Len(SequenceNr) <= 3 And Len(SequenceName) < 15 Then
               'Cells(1, 1).Select
               If oldseqNr <> SequenceNr Then
                  If SequenceNr = "6 , 4" Then
                     Debug.Print SequenceNr
                     End If
                     
                    ' Worksheets.Add After:=Worksheets(Worksheets.Count)
                     'Worksheets(Worksheets.Count).Name = SequenceNr
                     
                     Line Input #F, temp '
                     'Worksheets(SequenceNr).Activate
                     
                     For l = 1 To CntIntr
                        temp = Intrument(l)
                     Next l
                     'Debug.Print "Seq. Nr: " & SequenceNr
                     cntrow = 1
               End If
            
               oldseqNr = SequenceNr
            End If
        Case 4
            Input #F, d1    ' Total Stop Time
            Input #F, d2    ' Squence Time
        Case 5
            Line Input #F, temp ' Test Referance
            TestReferance = temp
            Line Input #F, temp ' Material
            'TestMaterial = temp
            temp = temp
        Case 6  ' Average Data
            For va_loop = 1 To 16
                Input #F, temp  ' Average Data
            Next va_loop
        Case 7
            Input #F, temp  ' Brake Time Secs
            Input #F, temp  ' Revs to Stop
        Case 8  ' Results Format
            Line Input #F, temp
        Case 10 ' Quality Values Info
            For l = 1 To 40
                Line Input #F, temp ' Quality Value Info
            Next l
        Case 11 ' Quality String Info
            For l = 1 To 6
                Line Input #F, temp ' Quality String Info
            Next l
        Case 12 ' ATE Info
            Input #F, temp  ' Servo Channel Number
            Input #F, temp  ' Servo Demand Value
        Case 13 ' MFDD + Extras
            Input #F, d1    ' MFDD
            For l = 1 To 5
                Input #F, d1
            Next l
        Case 20 ' Torque Offset ****(Not Used)****
            '# Line Input #f, temp
            '# Input #f, d1
            '# If c2 = 1 Then ' end 1
            '#    IOffset(Torque1) = d1
            '# Else ' end 2
            '#    IOffset(Torque2) = d1
            '# End If
        Case 21 ' thermocouple bits for Roulunds  ****(Not Used)****
            For l = 1 To 18
                Input #F, d1
            Next l
        Case 22 ' Running ISim
            Input #F, d1    ' Inertia Simulation 1 = On 0 = Off
        Case 23 ' Comments
            Line Input #F, temp ' Comments
        Case 24 ' Measurements
            For l = 1 To 17
                Input #F, d1    ' Measurements
            Next l
        Case 25 ' Measurements
            For l = 1 To 26
                Input #F, d1    ' Measurements
            Next l
        Case 26 ' Inertia
            Input #F, d1
        Case 27 ' Run Test Channel Information  ****(Not Used)****
            Input #F, d1, d1, d1, d1
        Case 28 ' Time and Date Info
            Input #F, d1
            Input #F, d1
            Input #F, d1
        Case 29 ' 6 Lines of Comments
            For l = 1 To 6
                Line Input #F, temp ' Comments
            Next l
        Case 30 ' sequence start date and time
            Line Input #F, temp ' number
            Line Input #F, temp ' name
            Line Input #F, temp ' date
            Line Input #F, temp ' time
        End Select
    Wend
    Close #F
    
    MsgBox "Import done...!"
End Sub

'###########################################################################
Sub DBTableNames()
    
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
    
End Sub
