VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Messenger"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkSavePassword 
      BackColor       =   &H00D3D3C3&
      Caption         =   "Save Password"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox cboServer 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "217.136.149.192"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogin.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel As Boolean

Private Sub cmdCancel_Click()
On Error Resume Next
Cancel = True
Hide
End Sub


Private Sub cmdOk_Click()
On Error Resume Next
SaveSetting "vasilakis Messenger", "Server", "Host", Me.cboServer.Text
Hide
End Sub

Private Sub Form_Activate()
'SetWindowPos hWnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Load()
On Error Resume Next
Cancel = False
txtUser.Text = GetSetting("vasilakis Messenger", "Login", "Email", "user@domain.com")
Me.cboServer.Text = GetSetting("vasilakis Messenger", "Server", "Host", "217.136.149.192")

chkSavePassword.Value = GetSetting("vasilakis Messenger", "Login", "SavePass", "0")
If chkSavePassword.Value = 1 Then
    txtPass.Text = GetSetting("vasilakis Messenger", "Login", "Password", "")
End If
End Sub


Private Sub txtPass_GotFocus()
txtPass.SelStart = 0
txtPass.SelLength = Len(txtPass.Text)
End Sub


Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = Len(txtUser.Text)
End Sub


