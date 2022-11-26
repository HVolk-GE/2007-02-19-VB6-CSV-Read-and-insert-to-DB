Attribute VB_Name = "A_Coments"
' Bug Report:
' Microface universal tools, vom 19.02.2007
' # vom 20.02.2007:
'####################################################################################
' Bevor die Daten aus den CSV (AV & AC) Dateien in die Datenbank importiert werden,
' muessen folgende dinge noch bearbeitet werden;
' 1.) Replace fuer ein "," muss ein "." gesetzt werden, -> DBUSDEF=Yes -> 20.02.2007
' 2.) Aus DATETIME und Time muss eine Spalte DATETIME gemacht werden -> 20.02.2007
'####################################################################################
' # vom 21.02.2007:
' In die Tabelle, die erstellt wurde, koennen keine Daten importiert
' werden ! Hat sich erledigt, wenn tabellennamen ein "-" beinhaltet
' dann muessen "`" & sTable & "`" gesetzt werden.
'
'
