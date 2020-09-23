VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Backup Utility using SQLDMO"
   ClientHeight    =   4275
   ClientLeft      =   3450
   ClientTop       =   1845
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7710
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   3960
      TabIndex        =   23
      Top             =   3060
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4260
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Database List "
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   3900
      TabIndex        =   19
      Top             =   780
      Width           =   3765
      Begin VB.ComboBox cmbDb 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   320
         Width           =   2445
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   380
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup"
      Height          =   375
      Left            =   5850
      TabIndex        =   18
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Frame fraPathDetails 
      Caption         =   "Path && File Information"
      ForeColor       =   &H00000080&
      Height          =   1605
      Left            =   3900
      TabIndex        =   15
      Top             =   1680
      Width           =   3765
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1230
         TabIndex        =   16
         Top             =   510
         Width           =   2385
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   570
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1230
      TabIndex        =   13
      Top             =   3870
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Co&nnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   30
      TabIndex        =   12
      Top             =   3870
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Authentication Details "
      ForeColor       =   &H00000080&
      Height          =   1875
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   3795
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Width           =   5535
      End
      Begin VB.OptionButton optSQLAuth 
         Caption         =   "Use S&QL Server authentication"
         Height          =   252
         Left            =   90
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optWinNTAuth 
         Caption         =   "Use Windows &NT authentication"
         Height          =   252
         Left            =   90
         TabIndex        =   7
         Top             =   300
         Width           =   3015
      End
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1440
         Width           =   2292
      End
      Begin VB.TextBox txtLoginName 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Top             =   1110
         Width           =   2292
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label lblLoginName 
         AutoSize        =   -1  'True
         Caption         =   "Login name:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1170
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Server Details "
      ForeColor       =   &H00000080&
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   3795
      Begin VB.CheckBox chkStartServer 
         Caption         =   "Start Server"
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   660
         Width           =   1425
      End
      Begin VB.ComboBox cmbServers 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Text            =   "cmbServers"
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Select the Server"
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By : Rohit Sharma"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3930
      TabIndex        =   22
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EEB68C&
      Caption         =   "SQL DB BACKUP UTILITY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   600
      TabIndex        =   14
      Top             =   0
      Width           =   7365
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmLogin.frx":1708A
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmLogin.frx":2E114
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSQLState As String
Dim WithEvents oBackupEvent As SQLDMO.Backup
Attribute oBackupEvent.VB_VarHelpID = -1
Dim gDatabaseName As String
Dim gBkupRstrFileName As String
Dim gBkupRstrFilePath As String

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 8/5/2004 18:13
' Author    : Rohit Sharma
' Purpose   : fill the combo with the available Servers on the network
'---------------------------------------------------------------------------------------

Private Sub Form_Load()
SQLServerList cmbServers
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmbServers_Click
' DateTime  : 8/5/2004 18:14
' Author    : Rohit Sharma
' Purpose   : selection of the server from the listed one
'---------------------------------------------------------------------------------------

Private Sub cmbServers_Click()
If Len(cmbServers.Text) > 0 Then
    cmdConnect.Enabled = True
Else
    cmdConnect.Enabled = False
End If

vSQLState = SQLServerStatus(cmbServers) 'call to the function to get the state of the function

If vSQLState = "Running" Then
    chkStartServer.Enabled = False
Else
    chkStartServer.Enabled = True
End If

End Sub

Private Sub cmbServers_Change()
cmbServers_Click
ServerName = cmbServers.Text
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdConnect_Click
' DateTime  : 8/5/2004 18:14
' Author    : Rohit Sharma
' Purpose   : connect the selected server
'---------------------------------------------------------------------------------------

Private Sub cmdConnect_Click()
On Error GoTo errhandle

'check if the server is running or not
'if stopped and if the start server check is marked then
'start the server

If chkStartServer.Value = 1 Then
    StartSQLServer cmbServers, txtLoginName.Text, txtPassword.Text
End If

'connect according to the authentication

If optSQLAuth.Value = True Then
    SqlConnect True, cmbServers.Text, txtLoginName.Text, txtPassword.Text
ElseIf optWinNTAuth.Value = True Then
    SqlConnect False, cmbServers.Text, txtLoginName.Text, txtPassword.Text
End If


'fill the database names for the server in the combo

FillDBList cmbDb

Exit Sub
errhandle:
    MsgBox Err.Description & vbCrLf & Err.Number, vbCritical, "Error"

End Sub

Private Sub cmbDb_Click()
txtFileName.Text = "c:\" & cmbDb.Text & ".bak"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdBackup_Click
' DateTime  : 8/5/2004 17:02
' Author    : Rohit Sharma
' Purpose   : to backup a sql database
'             the backup procedure could be called automatically
'             on the closure of the program
'---------------------------------------------------------------------------------------

Private Sub cmdBackup_Click()

ProgressBar1.Visible = True
On Error GoTo ErrHandler:
    
    Dim oBackup As SQLDMO.Backup
    
    gDatabaseName = cmbDb.Text
    Set oBackup = New SQLDMO.Backup
    Set oBackupEvent = oBackup ' enable events
    
    oBackup.Database = gDatabaseName
    gBkupRstrFileName = txtFileName.Text
    oBackup.Files = gBkupRstrFileName
    
    ' Delete the datafile to allow the application to create a brand new file.
    ' This will prevent attaching the new backup data to the old data if there
    ' is any.
    If Len(Dir(gBkupRstrFileName)) > 0 Then
        Kill (gBkupRstrFileName)
    End If
    
    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass
    
    ' Backup the database.
    oBackup.SQLBackup SqlState
    
    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault
   
    MsgBox "Backup complete", vbOKOnly + vbInformation, "Backup"
    
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
    
    Set oBackupEvent = Nothing ' disable events
    Set oBackup = Nothing
    
    Exit Sub

ErrHandler:
'    MsgBox "Error " & Err.Description
    Resume Next

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub optSQLAuth_Click()
txtLoginName.Enabled = True
txtPassword.Enabled = True

End Sub

Private Sub optWinNTAuth_Click()
txtLoginName.Enabled = False
txtPassword.Enabled = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : oBackupEvent_PercentComplete
' DateTime  : 8/5/2004 18:16
' Author    : Rohit Sharma
' Purpose   : shows the percent of task done
'---------------------------------------------------------------------------------------

Private Sub oBackupEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    ProgressBar1.Value = Percent
End Sub
