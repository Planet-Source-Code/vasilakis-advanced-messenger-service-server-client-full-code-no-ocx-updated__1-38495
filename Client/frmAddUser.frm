VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User to your list"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
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
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdduser 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the email address of the user you want to add. "
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAddUser.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblINfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   450
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel  As Boolean
Private Sub cmdAdduser_Click()
On Error Resume Next
If txtUser.Text = "" Then
    MsgBox "Please specify an email.", vbInformation, "Oops!"
    txtUser.SetFocus
    Exit Sub
End If
Cancel = False
Hide
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Cancel = True
Hide
End Sub


Private Sub Form_Activate()
SetWindowPos hwnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Load()
On Error Resume Next
Cancel = False
Me.Left = (Screen.Width / 2) - (Width / 2)
Me.Top = (Screen.Height / 2) - (Height / 2)
End Sub


