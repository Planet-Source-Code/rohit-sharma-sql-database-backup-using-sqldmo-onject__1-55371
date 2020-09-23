Attribute VB_Name = "modSQL"
'---------------------------------------------------------------------------------------
' Module    : modSQL
' DateTime  : 7/31/2004 12:14
' Author    : Rohit Sharma
' Purpose   : Module that handles all the possible SQL Operations
'---------------------------------------------------------------------------------------

Global SqlState As New SQLDMO.SQLServer
Global connSQL As New Connection
Global SqlTable As SQLDMO.Table
Global SqlDataBase As SQLDMO.Database
Global ServerName As String
Global SqlColumn As SQLDMO.Column


'---------------------------------------------------------------------------------------
' Procedure : StartSQLServer
' DateTime  : 7/31/2004 12:13
' Author    : Rohit Sharma
' Purpose   : Function to Start stopped Sql Server
'---------------------------------------------------------------------------------------

Public Function StartSQLServer(ServerName As String, Username As String, Password As String)
    
On Error GoTo BestHandler
    SqlState.Start False, ServerName, Username, Password
BestHandler:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SQLServerStatus
' DateTime  : 7/31/2004 12:13
' Author    : Rohit Sharma
' Purpose   : procedure to get the state of SQL server
'---------------------------------------------------------------------------------------

Public Function SQLServerStatus(SQLServerName) As Variant
On Error GoTo BestHandler

SqlState.Name = SQLServerName

If SqlState.Status = SQLDMOSvc_Running Then
    SQLServerStatus = "Running"
ElseIf SqlState.Status = SQLDMOSvc_Paused Then
    SQLServerStatus = "Paused"
ElseIf SqlState.Status = SQLDMOSvc_Stopped Then
    SQLServerStatus = "Stopped"
ElseIf SqlState.Status = SQLDMOSvc_Unknown Then
    SQLServerStatus = "Unknown"
ElseIf SqlState.Status = SQLDMOSvc_Continuing Then
    SQLServerStatus = "Continuing"
ElseIf SqlState.Status = SQLDMOSvc_Pausing Then
    SQLServerStatus = "Pausing"
ElseIf SqlState.Status = SQLDMOSvc_Starting Then
    SQLServerStatus = "Starting"
ElseIf SqlState.Status = SQLDMOSvc_Stopping Then
    SQLServerStatus = "Stopping"
End If


BestHandler:

ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'---------------------------------------------------------------------------------------
' Procedure : StopSQLServer
' DateTime  : 7/31/2004 12:13
' Author    : Rohit Sharma
' Purpose   : function to stop running server
'---------------------------------------------------------------------------------------

Public Function StopSQLServer(ServerName As String)
On Error GoTo BestHandler
SQLS.Name = ServerName
SQLS.Stop
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function


'---------------------------------------------------------------------------------------
' Procedure : Fill_cboSQLServerList
' DateTime  : 7/31/2004 12:12
' Author    : Rohit Sharma
' Purpose   : This procedure will scan the network and fill
'             the combo box with all the available servers.
'---------------------------------------------------------------------------------------

Sub SQLServerList(cmbSQLServerList As ComboBox)

Dim ServerList      As SQLDMO.NameList
Dim SQLApp          As SQLDMO.Application
Dim lngCounter      As Long

' Scan the network for SQL Servers and list them in the combo

Set SQLApp = New SQLDMO.Application
Set ServerList = SQLApp.ListAvailableSQLServers

cmbSQLServerList.Clear
'cmbSQLServerList.ItemData(cmbSQLServerList.NewIndex) = 1

For lngCounter = 1 To ServerList.Count
    cmbSQLServerList.AddItem ServerList.Item(lngCounter)
    cmbSQLServerList.ItemData(cmbSQLServerList.NewIndex) = lngCounter
Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FillDBList
' DateTime  : 7/31/2004 15:18
' Author    : Rohit Sharma
' Purpose   : procedure to fill list with available database in the server
'---------------------------------------------------------------------------------------

Public Sub FillDBList(cmb As ComboBox)
    
Set SqlState = New SQLDMO.SQLServer
SqlState.LoginSecure = True
SqlState.Connect ServerName
For Each SqlDataBase In SqlState.Databases
        cmb.AddItem SqlDataBase.Name
Next
    Exit Sub
End Sub

' Procedure : SqlConnect
' DateTime  : 8/4/2004 12:01
' Author    : Rohit Sharma
' Purpose   : procedure to connect to the selected SQL Server
'---------------------------------------------------------------------------------------

Public Sub SqlConnect(AuthOption As Boolean, SqlName As String, LoginName As String, Password As String)
If AuthOption = True Then 'for Sql autho
            Set SqlState = New SQLDMO.SQLServer
            SqlState.Connect SqlName, LoginName, Password
            txtLoginName.Enabled = False
            txtPassword.Enabled = False
        ElseIf AuthOption = False Then 'for winNT autho
            Set SqlState = New SQLDMO.SQLServer
            SqlState.LoginSecure = True
            SqlState.Connect SqlName
        End If
End Sub
