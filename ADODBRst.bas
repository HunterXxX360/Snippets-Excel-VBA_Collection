Dim strAccdb As String
Dim strConnStr As String
Dim adodbDB As New ADODB.Connection
Dim strTableName As String
Dim adodbRst As New ADODB.Recordset

strAccdb = "" 'Pfad der Access DB
strTableName = "" 'Tabellenname innerhalb der Access DB
'strConnStr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & strAccdb 'für .accdb
'strConnStr = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & strAccdb 'für .mdb

adodbDB.Open strConnStr
adodbRst.Open Source:=strTableName, ActiveConnection:=adodbDB ', CursorType:=adOpenStatic, LockType:=adLockReadOnly

'#################################
'#           CODE HERE           #
'#################################

adodbRst.Close
adodbDB.Close
