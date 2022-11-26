Attribute VB_Name = "CreateMySQL"
Public oConn As ADODB.Connection

Sub ConnectMySQL()
Dim oRS As Recordset
Dim sConn As String
Dim sServer As String
Dim sUserName As String
Dim sPassword As String
Dim sDBName As String

Set oConn = New ADODB.Connection
Set oRS = New ADODB.Recordset

oConn.Open "Provider=MSDASQL;DSN=joshuaDynoresults"

' Jetzt kann es losgehen...
' z.B. eine Abfrage erstellen...
'*oRS.Open "SELECT * FROM crack", oConn

'Sie können eine Verbindung zur MySQL-Datenbank aber _
 auch wie folgt herstellen:
' Connection-Objekt
' Instanzierung des Objekts
'# Set oConn = New ADODB.Connection

' Connection-String festlegen
' Server Hostname (oder IP)
'# sServer = DBHost '"localhost"

' Benutzerdaten
'# sUserName = DBUserName '"Benutzername"
'# sPassword = DBPassword '"Kennwort"

' Datenbank-Name
'# sDBName = dbName '"datenbankname"

'# sConn = "Provider=MSDASQL;Driver=MySQL ODBC 3.51 Driver;" & _
        "Server=" & sServer & ";Database=" & sDBName

' Connection öffnen
'# oConn.Open sConn, sUserName, sPassword
' Rest wie bei MS-Access Zugriff
End Sub

Sub CloseMySQLConn()
Dim oConn As ADODB.Connection
Dim oRS As Recordset
Dim sConn As String
Dim sServer As String
Dim sUserName As String
Dim sPassword As String
Dim sDBName As String

' Recordset erstellen
'Set oRS = New ADODB.Recordset

' Öffnen der ODBC Schnittstelle
' DSN ist der DataSourceName

'Sie können eine Verbindung zur MySQL-Datenbank aber auch wie folgt herstellen:

' Connection-Objekt

' Instanzierung des Objekts
'Set oConn = New ADODB.Connection
' Connection-String festlegen

oConn.Close

End Sub

Sub CreateMySQLFile()
Dim sZeilen0() As String, sZeilen1() As String, sZeilen2() As String
Dim sZeilen3() As String, lZeilen3 As Long, lZeilen0 As Long, lZeilen1 As Long
Dim lZeilen2 As Long, TestNr As Integer, TNr As String, TDNr As String
Dim VERSUCH As String, FileNames(4, 99) As String, FileNames0(99) As String
Dim j As Integer
 
    WriteZeile0 = 0
    j = 0
    i = Len(Form1.Combo1.Text)
 
    If i = 4 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 3)
    If i = 5 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 4)
 
    StrSoucreFile0 = SourcePath & Form1.Combo2.Text & "." & RESFile
 
    StrSoucreFile1 = Form1.Combo2.Text & StrSoucreExte1
    StrSoucreFile2 = SourcePath & StrSoucreFile1
    StrSoucreFile3 = StrDesFile2 & AVdataFile
    StrSoucreFile4 = StrDesFile2 & ACdataFile
  
    ' Hier ist es eine csv Datei:
    If i = 5 Then
         
         Form1.Command1.Enabled = False
         Form1.Command2.Enabled = False
         Form1.Combo1.Enabled = False
         Form1.Combo2.Enabled = False
         Form1.Text1.Enabled = False
         Form1.Text2.Enabled = False
         
        ' Lese hier nur die ersten 4 Zeilen in Array ein zum weiterenabarbeiten:
        Open StrSoucreFile0 For Input As 1
        
         For a = 0 To 3
             ReDim Preserve sZeilen3(lZeilen3 + 1) As String
             Line Input #1, sZeilen3(UBound(sZeilen3))
             lZeilen3 = UBound(sZeilen3)
         Next a
         
         VERSUCH = sZeilen3(lZeilen3)
         
        Close #1
        
        Open StrSoucreFile2 For Input As 1
            
        For a = 0 To 3
             ReDim Preserve sZeilen0(lZeilen0 + 1) As String
             Line Input #1, sZeilen0(UBound(sZeilen0))
             lZeilen0 = UBound(sZeilen0)
        Next a
        ' Close vorerst diese Datei spaeter wieder oeffnen um daten zu importieren
        Close #1

        ' Suche Testnummer und Tech. Daten nummer, leider statisch !
        StrTemp = ""
        For a = 0 To Len(sZeilen0(1))
            StrTemp = Mid(sZeilen0(1), a + 1, 1)
            If StrTemp = defSeperator Then '"," Then
               TNr = Mid(sZeilen0(2), 1, 6) ' Test-number
               TDNr = Mid(sZeilen0(2), 8, 6) ' Tech. Daten-number
               Exit For
            End If
        Next a
            
        If TestNr = 0 Then
           StrSoucreFile3 = StrSoucreFile3 & ".csv"
           StrSoucreFile4 = StrSoucreFile4 & ".csv"
        ElseIf TestNr = 1 Then
           StrSoucreFile3 = StrSoucreFile3 & "_" & TNr & ".csv"
           StrSoucreFile4 = StrSoucreFile4 & "_" & TNr & ".csv"
        End If
        
        d = 0
        z = 1
        
        ' Erste Spaltenueberschrift wird in Array geschrieben:
        For a = 0 To Len(sZeilen0(3)) + 1
            StrTemp = Mid(sZeilen0(3), a + 1, 1)
            If StrTemp = defSeperator And z = 1 Then
               FileNames(0, d) = Mid(sZeilen0(3), z, a)
               If FileNames(0, d) = "Sequence" Then
                  FileNames(0, d) = "ModulName"
               End If
               d = d + 1
               z = a + 1
             ElseIf StrTemp = defSeperator And z > 1 Then
               FileNames(0, d) = Mid(sZeilen0(3), z + 1, a - z)
               FileNames(0, d) = Trim(FileNames(0, d))
               d = d + 1
               z = a + 1
            End If
            
            If a = Len(sZeilen0(3)) Then
               FileNames(0, d) = Mid(sZeilen0(3), z + 1, a - z)
               FileNames(0, d) = Trim(FileNames(0, d))
               Exit For
             End If
        Next a
    
      e = 0
      z = 1
          ' Zweite Spaltenueberschrift wird in Array geschrieben:
          For a = 0 To Len(sZeilen0(4)) + 1
            StrTemp = Mid(sZeilen0(4), a + 1, 1)
            If StrTemp = defSeperator And z = 1 Then
               FileNames(1, e) = Mid(sZeilen0(4), z, a)
               e = e + 1
               z = a + 1
             ElseIf StrTemp = defSeperator And z > 1 Then
               FileNames(1, e) = Mid(sZeilen0(4), z + 1, a - z)
               FileNames(1, e) = Trim(FileNames(1, e))
               e = e + 1
               z = a + 1
            End If
            If a = Len(sZeilen0(4)) Then
               FileNames(1, e) = Mid(sZeilen0(4), z + 1, a - z)
               FileNames(1, e) = Trim(FileNames(1, e))
               Exit For
             End If
          Next a
          
          ' Aus beiden wird eine einzige Ueberschrift erstellt
          For Y = 0 To d - 1
              FileNames(2, Y) = FileNames(0, Y) & FileNames(1, Y) & """" & strDemilier
          Next Y
          
          FileNames(2, Y) = FileNames(0, Y) & FileNames(1, Y) & """"
          
          'Write AV Table here with Column Names
          Open StrSoucreFile3 For Output As 1
          Open StrSoucreFile2 For Input As 2
          
          For Y = 0 To d
              FileNames(3, 0) = FileNames(3, 0) & """" & FileNames(2, Y)
          Next Y
          
          FileNames(3, 0) = FileNames(3, 0) & strDemilier & """" & "SEQUENCE" & """"
          FileNames(3, 0) = FileNames(3, 0) & strDemilier & """" & "PRUEFLING" & """" & strDemilier
          FileNames(3, 0) = FileNames(3, 0) & """" & "SCHLUESSEL" & """" & strDemilier & """" & "VERSUCH" & """"
          
          Print #1, FileNames(3, 0)
          
          ' Open wieder daten tabelle um daten einzulesen und zu schreiben:
          ' Open StrSoucreFile2 For Input As 2
            While Not EOF(2)
                Y = 0
                ReDim Preserve sZeilen0(lZeilen0 + 1) As String
                Line Input #2, sZeilen0(UBound(sZeilen0))
                lZeilen0 = UBound(sZeilen0)
               
               If lZeilen0 > 8 Then
              
'##########################################################################################################

                  For a = 0 To Len(sZeilen0(lZeilen0))
                      StrTemp = Mid(sZeilen0(lZeilen0), a + 1, 1)
                      If StrTemp = defSeperator Then
                        If WriteZeile0 <> "" Then
                           If Len(WriteZeile0) > 1 Then
                             WriteZeile0 = WriteZeile0 & Mid(sZeilen0(lZeilen0), Y + 1, a - Y) & """" & strDemilier & """"
                           ElseIf Len(WriteZeile0) = 1 Then
                             WriteZeile0 = ""
                             WriteZeile0 = """" & Mid(sZeilen0(lZeilen0), Y + 1, a - Y) & """" & strDemilier & """"
                           End If
                             Y = a + 1
                             x = x + 1
                        Else
                             WriteZeile0 = """" & Mid(Trim(sZeilen0(lZeilen0)), 1, a) & """" & strDemilier & """"
                             Y = a + 1
                             x = x + 1
                        End If
                      End If
                  Next a

'##########################################################################################################
                 Y = 0
                  
                  For a = 0 To Len(WriteZeile0)
                      StrTemp = Mid(WriteZeile0, a + 1, 1)
                      If StrTemp = strdezmSep Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, Y + 1, a - Y) & ","
                         Y = a + 1
                         x = x + 1
                      ElseIf StrTemp = " " Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, Y + 1, a - Y)
                         Y = a + 1
                         x = x + 1
                      End If
                  Next a

'##########################################################################################################
                 WriteZeile = WriteZeile & Mid(sZeilen0(lZeilen0), Y + 1, a - Y) & """" & strDemilier & """"
                 Print #1, WriteZeile & "AV" & """" & strDemilier & _
                           """" & TDNr & """" & strDemilier & """" & TNr & """" & strDemilier & """" & VERSUCH; """"
                 
                 WriteZeile = ""
                 WriteZeile0 = ""
              
              End If
            
            Wend
            
          Close #1
          Close #2
         
         '#########################################################################
         '#
         '#  Jetzt wird ueberprueft ob es eine *.l* datei gibt, wenn ja wird auch
         '#  daraus die erste Spalte als ueberschrift gelesen und ...:
         '#
         '#########################################################################
         If Form1.File1.ListCount > 0 Then
         
                StrDesFile2 = SourcePath & Form1.File1.List(0)
  
                Open StrDesFile2 For Input As 1
    
                ' Lese hier nur die ersten 4 Zeilen in Array ein zum weiterenabarbeiten:
                For a = 0 To 3
                    ReDim Preserve sZeilen2(lZeilen2 + 1) As String
                    Line Input #1, sZeilen2(UBound(sZeilen2))
                    lZeilen2 = UBound(sZeilen2)
                    Exit For
                 Next a
            Close 1

'#########################################################################

            f = 0
            z = 1
           ' Erste Spaltenueberschrift wird in Array geschrieben:
            For a = 0 To Len(sZeilen2(1)) + 1
                StrTemp = Mid(sZeilen2(1), a + 1, 1)
                If StrTemp = defSeperator And z = 1 Then
                   FileNames0(f) = Mid(sZeilen2(1), z, a)
                   If FileNames0(f) = "Sequence" Then
                      FileNames0(f) = "STEPNR"
                   End If
                   f = f + 1
                   z = a + 1
                ElseIf StrTemp = defSeperator And z > 1 Then
                   FileNames0(f) = Mid(sZeilen2(1), z + 1, a - z)
                   FileNames0(f) = Trim(FileNames0(f))
                   'Stop Number
                   If FileNames0(f) = "Stop Number" Then
                      FileNames0(f) = "BRAKENR"
                   End If
                   'BR_TIME / Time
                   If FileNames0(f) = "Time" Then
                      FileNames0(f) = "BR_TIME"
                   End If
                   
                   If Len(FileNames0(f)) > 11 Then
                    For x = 0 To Len(FileNames0(f))
                          StrTemp0 = Mid(FileNames0(f), x + 1, 1)
                          If StrTemp0 = Chr(9) Then
                             FileNames0(f) = Mid(FileNames0(f), x + 2, Len(FileNames0(f)))
                             Exit For
                          End If
                      Next x
                    End If
                   
                   If Left(FileNames0(f), 4) = "Spee" Then
                      For x = 0 To Len(FileNames0(f))
                          StrTemp0 = Mid(FileNames0(f), x + 1, 1)
                          StrTemp1 = Mid(FileNames0(f), x + 2, 1)
                          If StrTemp0 = Chr(9) And StrTemp1 = Chr(9) Then
                             FileNames0(f) = Mid(FileNames0(f), x + 3, Len(FileNames0(f)))
                             Exit For
                          End If
                      Next x
                   End If
                   
                  ' DREHRI / Friction coefficient
                  ' If FileNames0(f) = "Friction coefficient" Then
                  '  FileNames0(f) = "STR2_1"
                  '    f = f
                  ' Else
                      f = f + 1
                  ' End If
                   
                      z = a + 1
                End If
                
                If a = Len(sZeilen2(1)) Then
                   FileNames0(f) = Mid(sZeilen2(1), z + 1, a - z)
                   FileNames0(f) = Trim(FileNames0(f))
                   For x = 0 To Len(FileNames0(f))
                          StrTemp0 = Mid(FileNames0(f), x + 1, 1)
                          StrTemp1 = Mid(FileNames0(f), x + 2, 1)
                          If StrTemp0 = Chr(9) And StrTemp1 = Chr(9) Then
                             FileNames0(f) = Mid(FileNames0(f), x + 3, Len(FileNames0(f)))
                             Exit For
                          End If
                   Next x
                   Exit For
                End If
            Next a
            
        For Y = 0 To f - 1
          If Y = 0 Then
             FileNames0(Y) = FileNames0(Y)
          End If
            FileNames0(Y) = FileNames0(Y) & """" & strDemilier
        Next Y
            FileNames0(Y) = FileNames0(Y) & """"
 
          Open StrSoucreFile4 For Output As 1
           
           For Y = 0 To f
                  FileNames(4, 0) = FileNames(4, 0) & """" & FileNames0(Y)
           Next Y
            
          FileNames(4, 0) = FileNames(4, 0) & strDemilier & """" & "SEQUENCE" & """"
          FileNames(4, 0) = FileNames(4, 0) & strDemilier & """" & "PRUEFLING" & """" & strDemilier
          FileNames(4, 0) = FileNames(4, 0) & """" & "SCHLUESSEL" & """" & strDemilier & """" & "VERSUCH" & """"
          
          Print #1, FileNames(4, 0)
          
          For i = 0 To Form1.File1.ListCount - 1
          
            StrDesFile2 = SourcePath & Form1.File1.List(i)
          
            Open StrDesFile2 For Input As 2
            
            lZeilen0 = 0
            
            While Not EOF(2)
                Y = 0
                ReDim Preserve sZeilen0(lZeilen0 + 1) As String
                Line Input #2, sZeilen0(UBound(sZeilen0))
                lZeilen0 = UBound(sZeilen0)
              
             If lZeilen0 > 1 Then
              
                  For a = 0 To Len(sZeilen0(lZeilen0))
                      StrTemp = Mid(sZeilen0(lZeilen0), a + 1, 1)
                      If StrTemp = defSeperator Then
                      If WriteZeile0 <> "" Then
                         WriteZeile0 = WriteZeile0 & Mid(sZeilen0(lZeilen0), Y + 1, a - Y) & """" & strDemilier & """"
                         Y = a + 1
                         x = x + 1
                         Else
                         WriteZeile0 = """" & Mid(sZeilen0(lZeilen0), 1, a) & """" & strDemilier & """"
                         Y = a + 1
                         x = x + 1
                       End If
                      End If
                  Next a
                  
                 WriteZeile0 = WriteZeile0 & Mid(sZeilen0(lZeilen0), Y + 1, a - Y) & """" & strDemilier & """"
                 
'###############################################################################
                 
                 Y = 0
                  
                  For a = 0 To Len(WriteZeile0)
                      StrTemp = Mid(WriteZeile0, a + 1, 1)
                      If StrTemp = strdezmSep Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, Y + 1, a - Y) & ","
                         Y = a + 1
                         x = x + 1
                      ElseIf StrTemp = " " Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, Y + 1, a - Y)
                         Y = a + 1
                         x = x + 1
                      End If
                  Next a
                  
'###############################################################################

                 WriteZeile = WriteZeile & Mid(WriteZeile0, Y + 1, a - Y)
                 
                 Print #1, WriteZeile & "AC" & """" & strDemilier & _
                           """" & TDNr & """" & strDemilier & """" & TNr & """" & strDemilier & """" & VERSUCH & """"
                 
                 WriteZeile = ""
                 WriteZeile0 = ""
              
              End If
            
            Wend
            
          Close #2
          Next i
          
          Close #1
          
        MsgBox "Create done, please you can find the data for :" & vbCr & _
               "AV Value in : " & StrSoucreFile3 & vbCr & _
               "AC Value in : " & StrSoucreFile4, vbInformation
          
          
    ElseIf Form1.File1.ListCount = 0 Then
       Open StrSoucreFile4 For Output As 1
       Print #1, ""
       Close #1
       
       MsgBox "Create done, please you can find the data for :" & vbCr & _
              "AV Value in : " & StrSoucreFile3, vbInformation

    End If
         
         Form1.Command2.Enabled = True
         Form1.Combo2.Enabled = True
         Form1.Text1.Enabled = True
         Form1.Text2.Enabled = True
         Form1.Combo2.Text = ""
         
         If Form1.Command1.Enabled = True Then
            Form1.Command1.Enabled = False
         End If
         
End If
End Sub
