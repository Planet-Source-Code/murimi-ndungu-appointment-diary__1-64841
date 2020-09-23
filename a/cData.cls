VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub OpenDB(dbName As String)
On Error GoTo ErrHandler:
Dim mPassword As String
Set objConn = New ADODB.Connection
Set objCmd = New ADODB.Command
Set objRS = New ADODB.Recordset
mPassword = "HarD24GeT$aS"
objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbName & ";Jet OLEDB:Database Password=" & mPassword
   
'strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'            "Data Source=" & dbName & _
'             "Jet OLEDB:Database Password=user"
   
        
objConn.Open
Exit Sub
ErrHandler:
Message = MsgBox(Err.Description, vbCritical)

Appstatus = "Invalid"
Exit Sub
End Sub

