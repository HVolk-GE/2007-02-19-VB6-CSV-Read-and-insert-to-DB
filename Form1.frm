VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Little-tools-farm Microface data export"
   ClientHeight    =   2625
   ClientLeft      =   4710
   ClientTop       =   3690
   ClientWidth     =   6540
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6540
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5640
      Top             =   120
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Create direct to DB"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Change only Charaters"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.FileListBox File2 
      Height          =   1065
      Left            =   8160
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   7560
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Close after copy"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   3000
      List            =   "Form1.frx":030C
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   240
      Picture         =   "Form1.frx":030E
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Little tools farm® H. Volk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4680
      TabIndex        =   11
      Top             =   2400
      Width           =   1770
   End
   Begin VB.Label Label5 
      Caption         =   "Target Path :"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Test Nr. :"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "File Type :"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Source Path :"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----==== sonstige Konstanten ====----
Private Const LB_DELETESTRING As Long = &H182

Private Declare Sub Sleep Lib "KERNEL32" _
        (ByVal dwMilliseconds As Long)

' ----==== USER32 API Deklarationen ====----
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Sub RemoveFiles(ByVal ctrlFileListBox As FileListBox, _
    ByVal sFiles As String)
    
    Dim lIndex As Long
    Dim lIndex1 As Long
    Dim sFilesArray() As String
    
    ' sind keine zu entfernenden Dateien
    ' angegeben, dann Sub verlassen
    If Len(sFiles) = 0 Then Exit Sub
    
    'Splitten der zu entfernenden Dateien
    sFilesArray = Split(sFiles, ";")
    
    With ctrlFileListBox
        
        ' Jedes Element (von hinten beginnend) durchgehen
        For lIndex = .ListCount - 1 To 0 Step -1
            
            ' Array durchgehen
            For lIndex1 = LBound(sFilesArray) _
            To UBound(sFilesArray)
                
                ' Element mit dem Muster vergleichen
                If LCase$(.List(lIndex)) Like _
                LCase$(sFilesArray(lIndex1)) Then
                    
                    ' Löschen des Elementes aus der
                    ' FileListBox
                    Call SendMessage(.hWnd, _
                    LB_DELETESTRING, lIndex, 0)
                    
                    ' Schleife verlassen
                    Exit For
                End If
            Next lIndex1
        Next lIndex
    End With
End Sub

Private Sub cmdRemove()
    ' bestimmte Dateien aus der FileListBox entfernen
    ' Me.Combo1.Text
    'Call RemoveFiles(Me.File2, "*.prj;*.il;*.RES;*.csv") '"vb*.e?e;*.??l;??c*.o*")
    Call RemoveFiles(Me.File2, "*.prj;*.il;*.RES;*.csv") '"vb*.e?e;*.??l;??c*.o*")
End Sub

Private Sub Combo2_Change()
    If UCase(Automat) <> "YES" Then
        IniTal
        MeChange
    End If
End Sub

Private Sub Combo2_Click()
    IniTal
    MeChange
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command5_Click()
    frmAbout.Show
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Form_Initialize()

    Me.Text1.Text = ""
    Me.Text2.Text = ""
    Me.Combo1.Text = ""
    Me.Combo2.Text = ""
    Me.Check1.Enabled = False
    Me.Command1.Caption = "Create"
    Me.Command1.Enabled = False
    Me.Command2.Caption = "Exit"
    Me.Command2.Enabled = True

  '#########
  If PicBMP <> "" Then
     Set Me.Image1.Picture = LoadPicture(PicBMP)
  End If

' #################################################################
    IfCheckActiv = 0
' #################################################################
    mkDB = 0
    IniTal ' Original ini File for create a csv file
    IniRead ' ini File for insert to database
    
    Me.File1.Path = Me.Text1.Text
    Me.File1.FileName = Me.Combo1.Text
    
    Director
    TestNr = 0
    Me.ProgressBar1.Visible = False
   
    If UCase(Automat) = "YES" Then
       Me.Visible = False
       ' Path zu Original Datei aus iniDatei
       Me.File1.Path = Me.Text1.Text
       Me.File1.FileName = "*.csv"
       
       For cntLena = 1 To Len(Form1.File1.List(0))
            txtempo = Mid(Form1.File1.List(0), cntLena, 1)
            If txtempo = "." Then
               FilesLenAuto = cntLena - 1
               FileNamesAuto = Left(Form1.File1.List(0), FilesLenAuto)
               Me.Combo2.Text = FileNamesAuto
               If ReadMFRES = "NO" Then
                  ExportDatas
               End If
               Exit For
            End If
       Next
       Me.Combo2.Text = FileNamesAuto
       Me.Combo2.Enabled = False
       Me.Text1.Enabled = False
       Me.Text2.Enabled = False
       Me.Check1.Enabled = False
       Me.Check2.Enabled = False
       Me.Check3.Enabled = False
    End If
    
    If ReadMFRES = "YES" Then
       ReadMicrofaceResults ("C:\Microwin\Results\D-1883.RES")
    End If
    
    If MFTOCSV = "YES" And MFTODB = "NO" Then
        Me.Label5.Enabled = True
        Me.Text2.Enabled = True
    ElseIf MFTODB = "YES" And MFTOCSV = "NO" Then
        Me.Text1.Enabled = False
        Me.Label5.Enabled = False
        Me.Text2.Enabled = False
    Else
        MsgBox "You must have one configuration for create a csv file or import csv file to database !"
    End If
    
End Sub

Private Sub Command1_Click()
'If ReadMFRES = "NO" Then
'CheckTablesNames

If MFTOCSV = "YES" And MFTODB = "NO" And MKBOTH = "NO" Then
   ExportDatas
ElseIf MFTODB = "YES" And MFTOCSV = "NO" And MKBOTH = "NO" Then
   TBLCreate = 0
   MakeDefaultCSVToDB
ElseIf MFTODB = "YES" Or MFTOCSV = "YES" Then
  If MKBOTH = "YES" Then
     ExportDatas
     StrDesFile2 = Form1.Text2.Text
     Form1.Text1.Text = StrDesFile2
     Form1.Text2.Text = "C:\Temp\"
     TBLCreate = 0
     MakeDefaultCSVToDB
  End If
Else
   MsgBox "You must have one configuration for create a csv file or import csv file to database !"
End If

Form1.Combo2.Text = ""

End Sub

Public Sub ExportDatas()
Dim cntLena, FilesLenAuto As Integer
Dim sli As Integer
Dim DateTimCnt0, DateTimCnt1 As Integer

cntcols1 = 0

If Me.Check3.Value = 1 Then
   mkDB = 1
End If

If Me.Check2.Value <> 1 Then

    Me.File1.Path = Me.Text1.Text
    
    Me.File1.FileName = Me.Combo1.Text
    
    Me.File2.Path = Me.Text1.Text
    Me.File2.FileName = Me.Combo2.Text & ".*" '"*.*"
    
    cmdRemove
    
    lZeilen0 = 0
    lZeilen3 = 0
    lZeilen2 = 0
    
    WriteZeile = ""
    WriteZeile0 = 0
    WriteZeile1 = ""
    WriteLine1 = ""
    
    j = 0
    i = Len(Form1.Combo1.Text)
 
    If i = 4 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 3)
    If i = 5 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 4)
 
    SourcePath = Form1.Text1.Text
    StrDesFile2 = Form1.Text2.Text
 
    StrSoucreFile0 = SourcePath & Form1.Combo2.Text & "." & ResFile
    
    If UCase(ReadMFRES) = "YES" Then
       ResFile = StrSoucreFile0
       Call ReadMicrofaceResults(ResFile)
       Exit Sub
    End If
    
    StrSoucreFile1 = Form1.Combo2.Text & StrSoucreExte1
    StrSoucreFile2 = SourcePath & StrSoucreFile1
    StrSoucreFile3 = StrDesFile2 & AVdataFile
    StrSoucreFile4 = StrDesFile2 & ACdataFile
  
    Me.ProgressBar1.Visible = True
    
    intProgress = Form1.File1.ListCount + 1
    
    intValueFiles = 100 / intProgress
    
    ' Hier ist es eine csv Datei:
    If i = 5 Then
         
         Form1.Command1.Enabled = False
         Form1.Command2.Enabled = False
         Form1.Combo1.Enabled = False
         Form1.Combo2.Enabled = False
         Form1.Text1.Enabled = False
         Form1.Text2.Enabled = False
         
        ' Lese hier nur die ersten, aus posprgnum in der ini datei oder die
        ' ersten 10 Zeilen in Array ein zum weiterenabarbeiten:
        Me.ProgressBar1.Value = intValueFiles * 1
        
        Open StrSoucreFile0 For Input As 1
         
        If posprgnum > 0 Then
           intnr = posprgnum
        Else
           intnr = 10
        End If
        
         For a = 0 To intnr
             ReDim Preserve sZeilen3(lZeilen3 + 1) As String
             Line Input #1, sZeilen3(UBound(sZeilen3))
             temptxt00 = Left(sZeilen3(lZeilen3), 2)
             VERSUCH = Left(sZeilen3(lZeilen3), 8)
             VERSUCH = UCase(Trim(VERSUCH))
             lZeilen3 = UBound(sZeilen3)
             
             '#################################################################
             ' Muss statisch sein, da es noch keinen Weg gibt dieses dynamisch
             ' zu aendern !
             '#################################################################
             
             temptxt00 = UCase(Trim(temptxt00))
             
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
                If MFTODB = "YES" And MFTOCSV = "NO" Then
                    sTable = VERSUCH
                    CheckTablesNames
                End If
'########################################################################################
                a = 0
                Exit For
             End If
         Next a
         
             If a <> 0 Then
                MsgBox "Have not found a Schedule Number for FM Schedules !" & vbCr & _
                       "Please, contact a developmer for changed !" & vbCr & _
                       "Programm go down now !", vbCritical
                       End
             End If

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
            If StrTemp = defSeperator Then
               TNr = Mid(sZeilen0(2), 1, 6) ' Test-number
               TDNr = Mid(sZeilen0(2), 8, 6) ' Tech. Daten-number
               Exit For
            End If
        Next a
        
        If TNr <> "" Then
           StrSoucreFile3 = StrSoucreFile3 & "_" & TNr & ".csv"
           StrSoucreFile4 = StrSoucreFile4 & "_" & TNr & ".csv"
        Else
           StrSoucreFile3 = StrSoucreFile3 & ".csv"
           StrSoucreFile4 = StrSoucreFile4 & ".csv"
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

               If FileNames(0, d) = "DATETIME" Then
                  DateTimCnt0 = d
               ElseIf UCase(FileNames(0, d)) = "TIME" Then
                  DateTimCnt1 = d
               End If
               
               d = d + 1
               z = a + 1
            End If
            
            If a = Len(sZeilen0(3)) Then
               FileNames(0, d) = Mid(sZeilen0(3), z + 1, a - z)
               FileNames(0, d) = Trim(FileNames(0, d))
               Exit For
             End If
        Next a
        
      intB0 = 1
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
          For y = 0 To d - 1
              ReDim Preserve csvColumname(y)
              If y = DateTimCnt0 Then
                 csvColumname(y) = UCase(Trim(FileNames(0, DateTimCnt0))) & strDemilier ' & FileNames(1, y)))
              ElseIf y = DateTimCnt1 Then
                 temp = UCase(Trim(FileNames(0, DateTimCnt1)))  ' & FileNames(1, y)))
              Else
                 csvColumname(y) = UCase(Trim(FileNames(0, y) & FileNames(1, y)))
              End If

              FileNames(2, y) = FileNames(0, y) & Trim(FileNames(1, y)) & strDemilier
          Next y
          
          cntcols1 = d - 1
          
          FileNames(2, y) = Trim(FileNames(0, y)) & Trim(FileNames(1, y))

          'Write AV Table here with Column Names
          If mkDB = 0 Then
            Open StrSoucreFile3 For Output As 1
            Open StrSoucreFile2 For Input As 2
          Else
            Open StrSoucreFile2 For Input As 2
          End If
          
           For y = 0 To d
               If y = DateTimCnt0 Then
                 FileNames(3, 0) = Trim(FileNames(3, 0)) & Trim(FileNames(2, DateTimCnt0))
               ElseIf y = DateTimCnt1 Then
                  temp = Trim(FileNames(3, 0)) & Trim(FileNames(2, DateTimCnt1))
               Else
                  FileNames(3, 0) = Trim(FileNames(3, 0)) & Trim(FileNames(2, y))
               End If
           Next y
          
          FileNames(3, 0) = FileNames(3, 0) & strDemilier & "SEQUENCE"
          FileNames(3, 0) = FileNames(3, 0) & strDemilier & "PRUEFLING" & strDemilier
          FileNames(3, 0) = FileNames(3, 0) & "SCHLUESSEL" & strDemilier & "VERSUCH"
          FileNames(3, 0) = Replace(FileNames(3, 0), " ", "")
          
          '# Database Connect for AV Values
          If mkDB = 0 Then
             Print #1, FileNames(3, 0)
          Else
             FileNames(3, 0) = Replace(FileNames(3, 0), ";", ", ")
             dbConnect
          End If
          
          x = 0
' Open wieder daten tabelle um daten einzulesen und zu schreiben:
          
          While Not EOF(2)
                y = 0
                x = 0
                
                ReDim Preserve sZeilen0(lZeilen0 + 1) As String
                Line Input #2, sZeilen0(UBound(sZeilen0))
                lZeilen0 = UBound(sZeilen0)
                
               If lZeilen0 > 8 Then
               
'##########################################################################################################

                  sZeilen0(lZeilen0) = Replace(sZeilen0(lZeilen0), defSeperator, strDemilier)
                  sZeilen0(lZeilen0) = Replace(sZeilen0(lZeilen0), strdezmSep, defSeperator)
                  sZeilen0(lZeilen0) = Replace(sZeilen0(lZeilen0), " ", "")
                  
                  For a = 0 To Len(sZeilen0(lZeilen0))
                      StrTemp = Mid(sZeilen0(lZeilen0), a + 1, 1)
                      If StrTemp = strDemilier Then 'defSeperator Then
                       If WriteZeile0 <> "" Then
                          If Len(WriteZeile0) > 1 Then
                             If x = DateTimCnt0 And DateTimCnt0 < DateTimCnt1 Then
                                WriteZeile0 = WriteZeile0 & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) ' & strDemilier
                                WriteZeile0 = Replace(WriteZeile0, "-", ".")
                             ElseIf x = DateTimCnt0 And DateTimCnt0 > DateTimCnt1 Then
                                WriteZeile0 = WriteZeile0 & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                                WriteZeile0 = Replace(WriteZeile0, "-", ".")
                             ElseIf x = DateTimCnt1 And DateTimCnt0 < DateTimCnt1 Then
                                WriteZeile0 = WriteZeile0 & " " & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                             ElseIf x = DateTimCnt1 And DateTimCnt0 > DateTimCnt1 Then
                                WriteZeile0 = WriteZeile0 & " " & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) '& strDemilier
                             Else
                                WriteZeile0 = WriteZeile0 & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                             End If
                          ElseIf Len(WriteZeile0) = 1 Then
                             WriteZeile0 = ""
                             If x = DateTimCnt0 And DateTimCnt0 < DateTimCnt1 Then
                                WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) ' & strDemilier
                                WriteZeile0 = Replace(WriteZeile0, "-", ".")
                             ElseIf x = DateTimCnt0 And DateTimCnt0 > DateTimCnt1 Then
                                WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                                WriteZeile0 = Replace(WriteZeile0, "-", ".")
                             ElseIf x = DateTimCnt1 And DateTimCnt0 < DateTimCnt1 Then
                                WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                             ElseIf x = DateTimCnt1 And DateTimCnt0 > DateTimCnt1 Then
                                WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) '& strDemilier
                             Else
                                WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), y + 1, a - y) & strDemilier
                             End If
                          End If
                         y = a + 1
                         x = x + 1
                         Else
                         WriteZeile0 = Mid(Trim(sZeilen0(lZeilen0)), 1, a) & strDemilier
                         y = a + 1
                         x = x + 1
                       End If
                      End If
                  Next a

                  WriteZeile0 = WriteZeile0 & Mid(Trim(sZeilen0(lZeilen0)), y + 1, a)

                  WriteZeile = WriteZeile0
                  
                  WriteZeile = WriteZeile & strDemilier & "AV" & strDemilier & TDNr & strDemilier & TNr & strDemilier & VERSUCH
                                 
                  '# Values for AV Value
                  If mkDB = 0 Then
                     Print #1, WriteZeile
                  Else
                     WriteZeile = Replace(WriteZeile, ",", ".")
                     WriteZeile = "('" & Replace(WriteZeile, ";", "', '") & "')"
                    FileNames1(rowcnt0) = WriteZeile
                    
                    If rowcnt0 = 98 Then
'                       DoEvents
                       Sleep 50
                       dbRowAdd
                       rowcnt0 = 0
                    End If
                  End If
                 
                 WriteZeile = ""
                 WriteZeile0 = ""
                 intB0 = 1
                 rowcnt0 = rowcnt0 + 1
               End If
            Wend

          If mkDB = 0 Then
             Close #1
             Close #2
          Else
             Close #2
                    If rowcnt0 <> 0 Then
                       Sleep 50
                       dbRowAdd
                       rowcnt0 = 0
                    End If
             dbDisconnect
          End If

'#########################################################################
'#
'#  Jetzt wird ueberprueft ob es eine *.l* datei gibt, wenn ja wird auch
'#  daraus die erste Spalte als ueberschrift gelesen und ...:
'#
'#########################################################################

     intValueFiles = 100 / intProgress
          
    If Form1.File2.ListCount > 0 Then

                StrDesFile2 = SourcePath & Form1.File2.List(0)
  
                Open StrDesFile2 For Input As 1
    
               ' Lese hier nur die ersten 4 Zeilen in Array ein zum weiterenabarbeiten:
                For a = 0 To 3
                    ReDim Preserve sZeilen2(lZeilen2 + 1) As String
                    Line Input #1, sZeilen2(UBound(sZeilen2))
                    lZeilen2 = UBound(sZeilen2)
                   ' Exit For
                 Next a
            Close 1

'#########################################################################

            F = 0
            z = 1
           ' Erste Spaltenueberschrift wird in Array geschrieben:
            For a = 0 To Len(sZeilen2(1)) + 1
                StrTemp = Mid(sZeilen2(1), a + 1, 1)
                If StrTemp = defSeperator And z = 1 Then
                   FileNames0(F) = Mid(sZeilen2(1), z, a)
                   If UCase(FileNames0(F)) = "SEQUENCE" Then
                      FileNames0(F) = "STEPNR"
                   End If
                   F = F + 1
                   z = a + 1
                ElseIf StrTemp = defSeperator And z > 1 Then
                   FileNames0(F) = Mid(sZeilen2(1), z + 1, a - z)
                   FileNames0(F) = Trim(FileNames0(F))
                   'Stop Number
                   If UCase(FileNames0(F)) = "STOP NUMBER" Then
                      FileNames0(F) = "LOOP_1"
                   End If
                   'BR_TIME / Time
                   If UCase(FileNames0(F)) = "TIME" Then
                      FileNames0(F) = "BR_TIME"
                   End If
                  ' DREHRI / Friction coefficient
                   If UCase(FileNames0(F)) = "FRICTION COEFFICIENT" Then
                      FileNames0(F) = "Frictioncoefficient"
                   End If
                   
                   If Len(FileNames0(F)) >= 10 Then
                      For x = 0 To Len(FileNames0(F))
                          StrTemp0 = Mid(FileNames0(F), x + 1, 1)
                          If StrTemp0 = Chr(9) Then
                                FileNames0(F) = Mid(FileNames0(F), x + 2, Len(FileNames0(F)))
                                'Debug.Print FileNames0(f)
                                Exit For
                          End If
                      Next x
                    End If
                   ' Speed / Force
                  If Left(Trim(FileNames0(F)), 4) = "Spee" Or Left(FileNames0(F), 4) = "Forc" Then
                     For x = 0 To Len(FileNames0(F))
                         StrTemp0 = Mid(FileNames0(F), x + 1, 1)
                         StrTemp1 = Mid(FileNames0(F), x + 2, 1)
                         If StrTemp0 = Chr(9) And StrTemp1 = Chr(9) Then
                            FileNames0(F) = Mid(FileNames0(F), x + 3, Len(FileNames0(F)))
                            Exit For
                         ElseIf StrTemp0 = Chr(9) And StrTemp1 = "" Then
                            FileNames0(F) = Right(FileNames0(F), 1)
                            Exit For
                         ElseIf StrTemp0 = Chr(9) And StrTemp1 <> "" Then
                            FileNames0(F) = Right(FileNames0(F), 1)
                            Exit For
                         End If
                     Next x
                   End If
                   F = F + 1
                   z = a + 1
                End If
                
                If a = Len(sZeilen2(1)) Then
                   FileNames0(F) = Mid(sZeilen2(1), z + 1, a - z)
                   FileNames0(F) = Trim(FileNames0(F))
                   For x = 0 To Len(FileNames0(F))
                          StrTemp0 = Mid(FileNames0(F), x + 1, 1)
                          StrTemp1 = Mid(FileNames0(F), x + 2, 1)
                          If StrTemp0 = Chr(9) And StrTemp1 = Chr(9) Then
                             FileNames0(F) = Mid(FileNames0(F), x + 3, Len(FileNames0(F)))
                             Exit For
                          End If
                   Next x
                   Exit For
                End If
            
            Next a
            
        For y = 0 To F - 1
        
          If y = 0 Then
             FileNames0(y) = Trim(FileNames0(y))
          End If
             FileNames0(y) = Trim(FileNames0(y)) & strDemilier
          Next y
            
          FileNames0(y) = FileNames0(y)
 
          If mkDB = 0 Then
             Open StrSoucreFile4 For Output As 1
          End If
           
           For y = 0 To F
                  FileNames(4, 0) = FileNames(4, 0) & FileNames0(y)
           Next y
                     
          FileNames(4, 0) = FileNames(4, 0) & strDemilier & "SEQUENCE"
          FileNames(4, 0) = FileNames(4, 0) & strDemilier & "PRUEFLING" & strDemilier
          FileNames(4, 0) = FileNames(4, 0) & "SCHLUESSEL" & strDemilier & "VERSUCH"
          
          'Debug.Print FileNames(4, 0)
          
          '# Database Connect for AC Values
          If mkDB = 0 Then
             Print #1, FileNames(4, 0)
          Else
          ' Hat nun die Spaltenueberschrift und macht connect to DB:
             dbConnect
             FileNames(3, 0) = Replace(FileNames(4, 0), ";", ", ")
          End If
          
          intVal = 0
          
          intValueFiles = 100 / Form1.File2.ListCount
          Me.ProgressBar1.Value = 0
          
          For i = 0 To Form1.File2.ListCount - 1
            
            intVal = i + 1
            
            If mkDB = 0 Then
               Me.ProgressBar1.Value = intValueFiles * intVal '/ 2
            End If
            StrDesFile2 = SourcePath & Form1.File2.List(i)
          
            lZeilen0 = 0
            ReDim Preserve sZeilen0(lZeilen0) As String
          
          Open StrDesFile2 For Input As 2

          While Not EOF(2)
              y = 0
              ReDim Preserve sZeilen0(lZeilen0 + 1) As String
              Line Input #2, sZeilen0(UBound(sZeilen0))
              lZeilen0 = UBound(sZeilen0)
                
              While lZeilen0 = 1 And Not EOF(2)

                Line Input #2, WriteLine1
                  For a = 0 To Len(WriteLine1)
                      StrTemp = Mid(WriteLine1, a + 1, 1)
                      If StrTemp = defSeperator Then
                      If WriteZeile0 <> "" Then
                         WriteZeile0 = WriteZeile0 & Mid(WriteLine1, y + 1, a - y) & strDemilier
                         y = a + 1
                         x = x + 1
                         Else
                         WriteZeile0 = Mid(WriteLine1, 1, a) & strDemilier
                         y = a + 1
                         x = x + 1
                       End If
                      End If
                  Next a
                  
                 WriteZeile0 = WriteZeile0 & Mid(WriteLine1, y + 1, a - y) & strDemilier
                 
'###############################################################################
                 
                 y = 0
                  
                  For a = 0 To Len(WriteZeile0)
                      StrTemp = Mid(WriteZeile0, a + 1, 1)
                      If StrTemp = strdezmSep Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, y + 1, a - y) & defSeperator
                         y = a + 1
                         x = x + 1
                      ElseIf StrTemp = " " Then
                         WriteZeile = WriteZeile & Mid(WriteZeile0, y + 1, a - y)
                         y = a + 1
                         x = x + 1
                      End If
                  Next a
                  
'###############################################################################

                 WriteZeile = WriteZeile & Mid(WriteZeile0, y + 1, a - y)
                 
                  '# Values for AV Value
                  If mkDB = 0 Then
                     Print #1, WriteZeile & "AC" & strDemilier & _
                           TDNr & strDemilier & TNr & strDemilier & VERSUCH
                  Else
                     WriteZeile = WriteZeile & "AC" & strDemilier & _
                                  TDNr & strDemilier & TNr & strDemilier & VERSUCH
                     WriteZeile = Replace(WriteZeile, ",", ".")
                     WriteZeile = "('" & Replace(WriteZeile, ";", "', '") & "')"
                    
                    FileNames1(rowcnt0) = WriteZeile
                    
                    If rowcnt0 = 98 Then
                       Sleep 50
                       dbRowAdd
                       rowcnt0 = 0
                    End If

                  End If
                 WriteZeile = ""
                 WriteZeile0 = ""
                 WriteZeile1 = ""
                 WriteLine1 = ""
                 
                 If mkDB <> 0 Then
                    rowcnt0 = rowcnt0 + 1
                 End If
                 
              Wend
          Wend
            Close #2
          Next i
          
          If mkDB = 0 Then
             Close #1
             Else
             If rowcnt0 <> 0 Then
                dbRowAdd
                Sleep 50
                rowcnt0 = 0
             End If
          End If

        Me.ProgressBar1.Value = 100
        
        If mkDB = 0 Then
            If MESAG = "YES" Then
               MsgBox "Create done, please you can find the data for :" & vbCr & _
                      "AV Value in : " & StrSoucreFile3 & vbCr & _
                      "AC Value in : " & StrSoucreFile4, vbInformation
            End If
        Else
           dbDisconnect
           MsgBox "Have done..." & vbCr & _
                  "You have not faild Messages view ?" & vbCr & _
                  "Then you have done, all data in the database !", vbInformation
        End If
        
    ElseIf Form1.File2.ListCount = 0 Then
        If mkDB = 0 Then
            ' Open StrSoucreFile4 For Output As 1
            ' Print #1, ""
            ' Close #1
            If MESAG = "YES" Then
               MsgBox "Create done, please you can find the data for :" & vbCr & _
                   "AV Value in : " & StrSoucreFile3, vbInformation
            End If
        End If
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

     Me.ProgressBar1.Value = 0
     Me.ProgressBar1.Visible = False

ElseIf Me.Check2.Value = 1 Then
     ChangeOnlyCharaters
End If
     
If Automat = "YES" Then
    'If MESAG = "YES" Then
       MsgBox "Data export, ..done !"
    'End If
    If mkDB <> 0 Then
       Kill (StrSoucreFile3)
       Kill (StrSoucreFile4)
       End
    End If
End If
     
End Sub

