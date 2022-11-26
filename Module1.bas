Attribute VB_Name = "Create_ini"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias _
                  "GetPrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpDefault As String, _
                  ByVal lpReturnedString As String, _
                  ByVal nSize As Long, _
                  ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias _
                  "WritePrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpString As Any, _
                  ByVal lpFileName As String) As Long
                  
Private INIPath As String
 
Public Property Get sPath() As String
    sPath = INIPath
End Property
 
Public Property Let sPath(ByVal NewValue As String)
    INIPath = NewValue
End Property
 
Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)
    WritePrivateProfileString Section, Key, sValue, INIPath
End Sub
 
Public Sub WriteValue(ByVal Section As String, ByVal Key As String, ByVal vValue As Variant)
    WriteString Section, Key, CStr(vValue)
End Sub
 
Public Function GetIniString(ByVal Section As String, ByVal Key As String, _
        Optional ByVal Default As String = "") As String
    
    Dim sTemp As String
 
    sTemp = String(256, 0)
    GetPrivateProfileString Section, Key, "", sTemp, Len(sTemp), INIPath
    If InStr(sTemp, Chr$(0)) Then
        sTemp = Left$(sTemp, InStr(sTemp, vbNullChar$) - 1)
    Else
        sTemp = Default
    End If
    
    GetIniString = sTemp
End Function
 
Public Function GetIniLong(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Long = -1) As Long
Dim sTemp As String
 
    sTemp = GetIniString(Section, Key, CStr(Default))
    If IsNumeric(sTemp) Then
        GetIniLong = CInt(sTemp)
    'Else
        'Evtl. Fehlermeldung ausgeben
    End If
End Function
 
Public Function GetIniBool(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Boolean = False) As Boolean
    GetIniBool = CBool(GetIniLong(Section, Key, CInt(Default)))
End Function
 
Sub IniTal()
Dim ININame As String, LastProj1 As String
   
   Form1.Combo1.Clear
   
   INIPath = App.Path ' "C:\"
   ININame = "\Create.ini"
   INIPath = INIPath & ININame
   
   LastProj1 = GetIniString("Path", "Soucre", INIPath)  '***
   Form1.Text1.Text = LastProj1
   
   LastProj1 = GetIniString("Path", "Dest", INIPath)  '***
   Form1.Text2.Text = LastProj1
   StrDesFile2 = LastProj1

   LastProj1 = GetIniString("Files", "First", INIPath)  '***
   Form1.Combo1.AddItem LastProj1
   Form1.Combo1.Text = LastProj1
   
   LastProj1 = GetIniString("Files", "First", INIPath)  '***
   Form1.Combo1.AddItem LastProj1
   
   LastProj1 = GetIniString("Project", "Extention", INIPath)  '***
   ResFile = LastProj1
   
   LastProj1 = GetIniString("Project", "PPGNum", INIPath)  '***
   posprgnum = LastProj1
   
   LastProj1 = GetIniString("Config", "Country", INIPath)  '***
   strDemilier = LastProj1
   
   LastProj1 = GetIniString("Config", "MFTOCSV", INIPath)  '***
   MFTOCSV = UCase(LastProj1)
   
   LastProj1 = GetIniString("Config", "MFTODB", INIPath)  '***
   MFTODB = UCase(LastProj1)
   
   LastProj1 = GetIniString("Config", "Make_Both", INIPath) '***
   MKBOTH = UCase(LastProj1)
   
   LastProj1 = GetIniString("Config", "Messages", INIPath) '***
   MESAG = UCase(LastProj1)
      
   PicBMP = GetIniString("Config", "Logo", INIPath) '***
   PicBMP = App.Path & "\" & PicBMP
      
'# Checkboxes config:
   LastProj1 = GetIniString("Config", "CreateDBVisable", INIPath)  '***
   CreateDBVisable = LastProj1
   
   If UCase(CreateDBVisable) = "YES" Then
      Form1.Check3.Visible = True
      Else
      Form1.Check3.Visible = False
   End If
   
   LastProj1 = GetIniString("Config", "CreateDBEnable", INIPath)  '***
   CreateDBEnable = LastProj1
   
   If UCase(CreateDBEnable) = "YES" Then
      Form1.Check3.Enabled = True
      Else
      Form1.Check3.Enabled = False
   End If
   
   LastProj1 = GetIniString("Config", "CreateDBdef", INIPath)  '***
   CreateDB = UCase(LastProj1)
   
    If UCase(CreateDB) = "YES" Then
       Form1.Check3.Value = 1
       Else
       Form1.Check3.Value = 0
    End If
       
   LastProj1 = GetIniString("Config", "ChangeCharVisable", INIPath)  '***
   ChangeCharVisable = LastProj1
   
   If UCase(ChangeCharVisable) = "YES" Then
      Form1.Check2.Visible = True
      Else
      Form1.Check2.Visible = False
   End If
   
   LastProj1 = GetIniString("Config", "ChangeCharEnable", INIPath)  '***
   ChangeCharEnable = LastProj1
   
   If UCase(ChangeCharEnable) = "YES" Then
      Form1.Check2.Enabled = True
      Else
      Form1.Check2.Enabled = False
   End If
   
   LastProj1 = GetIniString("Config", "ChangeChardef", INIPath)  '***
   ChangeChardef = LastProj1
   
    If UCase(ChangeChardef) = "YES" Then
       Form1.Check2.Value = 1
       Else
       Form1.Check2.Value = 0
    End If
   
   LastProj1 = GetIniString("Config", "CloseAfCopyVisable", INIPath)  '***
   CloseAfCopyVisable = LastProj1
   
   If UCase(CloseAfCopyVisable) = "YES" Then
      Form1.Check1.Visible = True
      Else
      Form1.Check1.Visible = False
   End If
   
   LastProj1 = GetIniString("Config", "CloseAfCopyEnable", INIPath)  '***
   CloseAfCopyEnable = LastProj1

   If UCase(CloseAfCopyEnable) = "YES" Then
      Form1.Check1.Enabled = True
      Else
      Form1.Check1.Enabled = False
   End If
      
   LastProj1 = GetIniString("Config", "CloseAfCopyDef", INIPath)  '***
   CloseAfCopyDef = LastProj1

    If UCase(CloseAfCopyDef) = "YES" Then
       Form1.Check1.Value = 1
       Else
       Form1.Check1.Value = 0
    End If

   LastProj1 = GetIniString("Config", "Automat", INIPath)  '***
   Automat = UCase(LastProj1)
   
   '# Read only the RES-File:
   LastProj1 = GetIniString("Config", "ReadMFRES", INIPath)  '***
   ReadMFRES = UCase(LastProj1)

   LastProj1 = GetIniString("Config", "defSep", INIPath)  '***
   defSeperator = LastProj1
   sOldSeparator = LastProj1
   
   LastProj1 = GetIniString("Config", "dezmSep", INIPath)  '***
   strdezmSep = LastProj1
   sOldDecimal = LastProj1
   
   LastProj1 = GetIniString("Config", "defSep", INIPath)  '***
   sNewDecimal = LastProj1
   
   If sNewDecimal = "" Then sNewDecimal = GetEntry(&HE)
   
   LastProj1 = GetIniString("Config", "NewSeparator", INIPath)  '***
   sNewSeparator = LastProj1

   If sNewSeparator = "" Then sNewSeparator = GetEntry(&HC)

'   LastProj1 = GetIniString("data", "datapath", INIPath)  '***
    datapath = ""
   
   LastProj1 = GetIniString("dbFiles", "AVDB", INIPath)  '***
   AVdataFile = LastProj1
   
   LastProj1 = GetIniString("dbFiles", "ACDB", INIPath)  '***
   ACdataFile = LastProj1
   
   Do While strDemilier = ""
            strDemilier = InputBox("Please insert here you csv field seperator !" & vbCr & _
                                   "You can insert the countrycode (US - USA, UK-England and DE-German) in ini File" & vbCr & _
                                   "For default in application Path !", "Seperator failed !")

   Loop
   
         If strDemilier = "DE" Then
            strDemilier = ";"
         ElseIf strDemilier = "UK" Then
            strDemilier = ","
         ElseIf strDemilier = "US" Then
            strDemilier = ","
         End If
End Sub

Sub Director()
Dim i As Integer, a As Integer, SourcefileName As String
Dim tmpName As String

   For i = 0 To Form1.File1.ListCount
      
      SourcefileName = Form1.File1.List(i)
      
      For a = 1 To Len(SourcefileName)
          tmpName = Mid(SourcefileName, a, 1)
          If tmpName = "." Then Exit For
      Next a
      
      tmpName = Left(SourcefileName, a - 1)
      Form1.Combo2.AddItem tmpName
   Next i

End Sub

Sub MeChange()
 
 If Form1.Combo1.Text <> "" Then
     Form1.Command1.Enabled = True
 End If
  
  SourcePath = Form1.Text1.Text
  DestiPath = Form1.Text2.Text
  Form1.File1.FileName = SourcePath & Form1.Combo2.Text & ".l*"

End Sub

Sub ReadFile()
Dim rs As New Recordset
 
     StrSoucreFile1 = Form1.Combo2.Text & StrSoucreExte1
     
     For i = 0 To Form1.File1.ListCount
         ArrayCopyFileNam(i) = Form1.File1.List(i)
     Next i
     
     a = i - 1
     StrSoucreFile2 = SourcePath & StrSoucreFile1
     StrDesFile2 = DestiPath & StrSoucreFile1
          
     If a >= 0 And SourcePath <> "" And DestiPath <> "" Then
       If a > 0 Then
        For i = 0 To a
           If ArrayCopyFileNam(i) <> "" Then
           StrSoucreFile2 = SourcePath & ArrayCopyFileNam(i)
           StrDesFile2 = DestiPath & ArrayCopyFileNam(i) & ".csv"
           FileCopy StrSoucreFile2, StrDesFile2
           Else
           Exit For
           End If
        Next i
       End If
     End If

End Sub

