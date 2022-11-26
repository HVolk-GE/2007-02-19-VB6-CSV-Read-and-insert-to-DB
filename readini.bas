Attribute VB_Name = "config_ini"
'#################################################################################
'### For read the ini-files
'#################################################################################
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
                  
Public Property Get sPath() As String
    sPath = INIPath1
End Property
 
Public Property Let sPath(ByVal NewValue As String)
    INIPath1 = NewValue
End Property
 
Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)
    WritePrivateProfileString Section, Key, sValue, INIPath1
End Sub
 
Public Sub WriteValue(ByVal Section As String, ByVal Key As String, ByVal vValue As Variant)
    WriteString Section, Key, CStr(vValue)
End Sub
 
Public Function GetIniString(ByVal Section As String, ByVal Key As String, _
        Optional ByVal Default As String = "") As String
        Dim sTemp As String
 
    sTemp = String(256, 0)
    GetPrivateProfileString Section, Key, "", sTemp, Len(sTemp), INIPath1
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
 
Public Sub IniRead()

INIPath1 = App.Path & "\config.ini"

sServer = GetIniString("dbconnect", "servername", INIPath1)
sUsername = GetIniString("dbconnect", "username", INIPath1)
sPassword = GetIniString("dbconnect", "password", INIPath1)
sDBName = GetIniString("dbconnect", "dbname", INIPath1)
sTable = GetIniString("dbconnect", "tablename", INIPath1)

DBUSDEF = GetIniString("dbconnect", "DBUSDEF", INIPath1)

DBUSDEF = UCase(DBUSDEF)

i = 0
Masch = "cols" & i
Values0 = GetIniString("dbcolumns", Masch, INIPath1)
Values0 = UCase(Trim(Values0))
Do While Values0 <> ""
   If Values0 <> "" Then
        ReDim Preserve DbColumname(i)
        ReDim Preserve CreatTableCols(i)
        Values0 = UCase(Trim(Values0))
        If Values0 = "BETRIEB" Or Values0 = "MFDATETIME" Or _
           Values0 = "DREHRI" Or Values0 = "MODULNAME" Or _
           Values0 = "PRUEFLING" Or Values0 = "PRUEFSTAND" Or _
           Values0 = "SCHLUESSEL" Or Values0 = "SEQUENCE" Or _
           Values0 = "VERSUCH" Then
           CreatTableCols(i) = Values0 '& " varchar(25) collate latin1_general_ci default '0', "
        Else
           CreatTableCols(i) = Values0 '& " double default '0', "
        End If
        DbColumname(i) = UCase(Trim(Values0))
        gcntcols = i
   Else
        Exit Do
   End If
   i = i + 1
   Masch = "cols" & i
   Values0 = GetIniString("dbcolumns", Masch, INIPath1)
Loop
 
End Sub

Sub IniWrite()
'WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)

End Sub
