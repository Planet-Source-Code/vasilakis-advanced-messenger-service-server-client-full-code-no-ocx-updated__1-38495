VERSION 5.00
Begin VB.Form frmAuth 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Authorization"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAdd 
      BackColor       =   &H00D3D3C3&
      Caption         =   "Add this user to my list."
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chkAllow 
      BackColor       =   &H00D3D3C3&
      Caption         =   "I allow this user to view my online/offline status, and send me messages."
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      Caption         =   "User email@email.com has added you to his/her list."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      TabIndex        =   0
      Top             =   300
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAuth.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
If chkAllow.Value = 0 Then
    frmMain.SendData "allowlist;" & GetPiece(Tag, vbLf, 2) & vbLf & "0"
Else
    frmMain.SendData "allowlist;" & GetPiece(Tag, vbLf, 2) & vbLf & "1"
End If
If chkAdd.Enabled = True Then
    If chkAdd.Value = 1 Then
        frmMain.SendData "adduser;" & LCase$(GetPiece(Tag, vbLf, 2))
    End If
End If
Unload Me
End Sub


Private Sub Form_Activate()
SetWindowPos hWnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

