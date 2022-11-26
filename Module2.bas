Attribute VB_Name = "MySQLdb"
Option Explicit

' Fehlerausgabe bei Verbindungsfehler
Private Function MySQL_Error() As Boolean
  With oConn.Error
    If .Number = 0 Then Exit Function
        MsgBox "Error " & .Number & ": " & .Description
        MySQL_Error = True
        End
  End With
End Function

Public Sub dbRowAdd()
' Datensatz hinzufügen
Dim nResult As Long
Dim sVorname, sNachname As String 'Vorsicht hat sich geaendert !
Dim sTestTyp, sProblem, sMaschine, sResult, sDescription, sTestnumber As String
  i = 0
  'If Not oConn Is Nothing Then oConn.CloseConnection
  If MySQL_Error() = False Then
     'Set oConn = Nothing
      ' Datensatz in Tabelle einfügen
      For i = 0 To rowcnt0 - 1
          'DoEvents
          oConn.Execute "INSERT INTO " & sTable & " (" & FileNames(3, 0) & ") " & _
                        "VALUES " & FileNames1(i)
      Next
  End If

End Sub

Public Sub dbCreateTable()
Dim sSQL, bSQL As String
  
 ' Tabelle erstellen
   sSQL = "CREATE TABLE IF NOT EXISTS " & sTable & " (`idx` BIGINT(20) NOT NULL AUTO_INCREMENT PRIMARY KEY) ENGINE = MYISAM ;"
  
  oConn.Execute sSQL
  
  For i = 0 To gcntcols
      bSQL = "ALTER TABLE " & sTable & " ADD " & CreatTableCols(i) '& " VARCHAR( 20 ) NULL DEFAULT '0';"
      oConn.Execute bSQL
  Next

  If MySQL_Error() = False Then
     If MESAG = "YES" Then
        MsgBox "Tabelle " & sTable & " wurde erstellt."
     End If
  End If
End Sub

Public Sub dbDisconnect()
  ' Verbindung beenden
  If Not oConn Is Nothing Then oConn.CloseConnection
  If MySQL_Error() = False Then
    Set oConn = Nothing
  End If
End Sub

' Tabelle löschen
Public Sub dbDropTable()
  Dim sSQL As String
  
  If MsgBox("Tabelle wirklich löschen?", vbYesNo, "Löschen") = vbYes Then
    ' Tabelle löschen
    sSQL = "DROP TABLE IF EXISTS " & sTable
    oConn.Execute sSQL
    If MySQL_Error() = False Then
    
    End If
  End If
End Sub

Public Sub dbQueryTable()

  ' RS-Objekt vom Connection-Objekt ableiten und
  ' Status anzeigen.
  '
  ' Nicht wie bei ADO. Bei der MyVbQl.Dll wird das
  ' Recordset direkt von der Connection abgeleitet.
  
  Dim sSQL As String
  Dim bError As Boolean
  
  ' Alle Datensätze selektieren
  
  sSQL = "SELECT * FROM " & sTable
  Set oRs = oConn.Execute(sSQL)
  bError = MySQL_Error()
 ' Recordset schließen
  If Not oRs Is Nothing Then oRs.CloseRecordset
  Set oRs = Nothing
End Sub

Public Sub dbConnect()
  ' Wir öffnen die Verbindung zum MySQL Server
  ' Statt 'Localhost' kann auch die IP verwendet werden. Diese
  ' erfahren Sie im WinMySQLAdmin im Register'Environment',
  ' wenn Sie auf 'Extendet Server Status' klicken.

  oConn.OpenConnection sServer, _
    sUsername, sPassword, sDBName

  ' Statusabfrage
  If (oConn.State = MY_CONN_CLOSED) Then
    ' Falls Verbindung nicht geöffnet, Fehlerangabe!
    MySQL_Error
  Else
    ' Bei erfolgreicher Verbindung, Verbindungsdaten ausgeben
    'MsgBox "Connected to Database: " & oConn.dbName, _
      vbInformation, "MySQL-Dyno-Results-Database"
      
    ' Prüfen, ob Tabelle existiert
    'Set oRs = oConn.Execute("SELECT * FROM " & sTable)

    Set oRs = oConn.Execute("SELECT * FROM " & sTable)
   
    ' Recordset sschließen
    If Not oRs Is Nothing Then oRs.CloseRecordset
    
    Set oRs = Nothing
    
    End If
End Sub

Public Sub CheckTablesNames()
'SHOW TABLES
Dim sSQL As String
Dim i As Integer

oConn.CloseConnection

dbConnect

sSQL = "SHOW TABLES"
Set oRs = oConn.Execute(sSQL)

If Not oRs Is Nothing Then

    oRs.MoveFirst

    Do While Not oRs.EOF And TBLCreate <> 1
        If sTable = "`" & oRs.Fields(0).Value & "`" Then
            TBLCreate = 1
            Exit Do
        End If
        oRs.MoveNext
    Loop

    If TBLCreate = 0 Then
       dbCreateTable
       dbQueryTable
       TBLCreate = 1
    End If

    If Not oRs Is Nothing Then oRs.CloseRecordset

    Set oRs = Nothing

    If Not oConn Is Nothing Then oConn.CloseConnection
    
End If
    
End Sub

' Standard-Browser starten und WWW-Seite aufrufen
Public Sub URLGoTo(ByVal hWnd As Long, ByVal URL As String)
  Screen.MousePointer = 11
  If Left$(URL, 7) <> "http://" Then URL = "http://" & URL
  Call ShellExecute(hWnd, "Open", URL, "", "", 3)
  Screen.MousePointer = 0
End Sub

