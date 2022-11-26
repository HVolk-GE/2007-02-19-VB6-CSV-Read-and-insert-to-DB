Attribute VB_Name = "GlobalVariables"
Option Explicit
'#################################################################################
'*** Public Variablen
'#################################################################################
Public Const URL As String = "www.little-tools-farm.de"
Public IfCheckActiv As Integer
Public Masch, INIPath1, ReadMFRES As String
' Servername und Benutzerdaten
Public sServer, sUsername, sPassword, sDBName, sTable, temptxt00 As String
Public DbColumname(), CreatTableCols(), csvColumname(), DBWriteZeile As String
Public gcntcols, cntcols1, mkDB, rowcnt0 As Integer
Public MFTOCSV, MFTODB, MKBOTH, MESAG, DBUSDEF, PicBMP As String
Public TBLCreate As Integer

' Checkbox, Button and ComboBox values visable or enables:
Public CreateDBEnable, CreateDBVisable, CreateDB As String
Public ChangeCharVisable, ChangeCharEnable, ChangeChardef As String
Public CloseAfCopyVisable, CloseAfCopyEnable, CloseAfCopyDef As String
Public Automat, FileNamesAuto As String

'#####################################################################
Public StrSoucreFile1 As String, StrSoucreExte1 As String
Public SourcePath As String, DestiPath As String, StrSoucreFile0 As String
Public StrSoucreFile2 As String, StrDesFile2 As String, ResFile As String
Public i As Integer, a As Integer, b As String, c As String
Public intb As Integer, inta As Integer, intc As Integer
Public ArrayCopyFileNam(99) As String, d As Integer, e As Integer, F As String
Public datapath As String, strdezmSep As String
Public AVdataFile As String, ACdataFile As String, strDemilier As String
Public defSeperator As String, posprgnum As Integer

Public sZeilen0() As String, sZeilen1() As String, sZeilen2() As String
Public sZeilen3() As String, lZeilen3 As Long, lZeilen0 As Long, lZeilen1 As Long
Public lZeilen2 As Long, TestNr As Integer, TNr As String, TDNr As String
Public VERSUCH As String, FileNames(4, 99) As String, FileNames0(99), FileNames1(99) As String
Public FileNames2()
Public WriteLine1 As String, WriteZeile1, WriteZeile As String
Public intProgress As Integer, intValueFiles As Integer
Public intVal As Integer
Public sOldSeparator, sOldDecimal, sNewDecimal, sNewSeparator As String
Public DBHost As String, DBUserName As String, DBPassword As String, dbName As String

' Die Deklarationen sind fast wie bei ADO
' Connection-Object
Public oConn As New MYSQL_CONNECTION

' Recordset-Object
Public oRs As MYSQL_RS

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

