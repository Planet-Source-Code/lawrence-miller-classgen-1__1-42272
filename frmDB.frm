VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDB 
   Caption         =   "Select Database"
   ClientHeight    =   2895
   ClientLeft      =   4830
   ClientTop       =   3855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   2865
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4635
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   585
         Left            =   1950
         TabIndex        =   13
         Top             =   2040
         Width           =   1035
      End
      Begin MSComDlg.CommonDialog cdOpen 
         Left            =   270
         Top             =   2220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   4260
         TabIndex        =   12
         ToolTipText     =   " Browse For Database "
         Top             =   930
         Width           =   315
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   1560
         Width           =   2715
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1530
         TabIndex        =   10
         Top             =   1230
         Width           =   2715
      End
      Begin VB.TextBox txtDB 
         Height          =   285
         Left            =   1530
         TabIndex        =   9
         Top             =   900
         Width           =   2715
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1530
         TabIndex        =   8
         Top             =   570
         Width           =   2715
      End
      Begin VB.OptionButton optDB 
         Caption         =   "Microsoft Access"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   7
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton optDB 
         Caption         =   "SQL Server"
         Height          =   195
         Index           =   0
         Left            =   1530
         TabIndex        =   6
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database Type:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   1230
         Width           =   1245
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   600
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    cdOpen.DialogTitle = "Select Database"
    cdOpen.InitDir = ApPath
    cdOpen.DefaultExt = ".mdb"
    cdOpen.Filter = "Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"
    cdOpen.FilterIndex = 0
    cdOpen.ShowOpen
    If cdOpen.FileName <> "" Then txtDB.Text = cdOpen.FileName
End Sub

Private Sub cmdOK_Click()
  Dim lResult As Long
  
    If txtDB.Text = "" Then
      MsgBox "You must enter a database name"
      txtDB.SetFocus
    End If
    
    If optDB(0).Value = True Then
      DBType = 1
    Else
      DBType = 2
    End If
    
    If DBType = 2 Then
      If Dir$(txtDB.Text, vbNormal) = "" Then
        MsgBox "Database not found"
        txtDB.SetFocus
        Exit Sub
      Else
        txtServer.Text = "MS Access 2000"
      End If
    End If
    
    If txtServer.Text = "" Then
      MsgBox "You must enter a server name"
      txtServer.SetFocus
      Exit Sub
    End If
      
    Server = Trim$(txtServer.Text)
    DBName = Trim$(txtDB.Text)
    User = Trim$(txtUser.Text)
    Password = Trim$(txtPassword.Text)
    
    lResult = PutString("ClassGen", "DatabaseType", CStr(DBType))
    lResult = PutString("ClassGen", "ServerName", Server)
    lResult = PutString("ClassGen", "DatabaseName", DBName)
    lResult = PutString("ClassGen", "User", User)
    lResult = PutString("ClassGen", "Password", Password)
    
    If OpenDataBase Then Unload Me
End Sub

Private Sub Form_Load()
    If DBType <> 0 Then
      optDB(DBType - 1).Value = True
    Else
      optDB(1).Value = True
    End If
      
    If Server <> "" Then txtServer.Text = Server
    If DBName <> "" Then txtDB.Text = DBName
    If User <> "" Then txtUser.Text = User
    If Password <> "" Then txtPassword.Text = Password
    
End Sub

Private Sub optDB_Click(Index As Integer)
    Select Case Index
      Case 0
        txtServer.Text = ""
        cmdBrowse.Visible = False
      Case 1
        txtServer.Text = "MS Access 2000"
        cmdBrowse.Visible = True
      End Select
End Sub
