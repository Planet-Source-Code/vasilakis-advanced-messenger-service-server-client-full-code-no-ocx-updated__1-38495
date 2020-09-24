VERSION 5.00
Begin VB.Form frmUserAdded 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add User"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ClipControls    =   0   'False
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
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUserAdded.frx":0000
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3945
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User has been added to your list."
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
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   240
      Picture         =   "frmUserAdded.frx":0093
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmUserAdded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub


Private Sub Form_Activate()
SetWindowPos hwnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

