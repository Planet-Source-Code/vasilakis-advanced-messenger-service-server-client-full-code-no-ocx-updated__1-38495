VERSION 5.00
Begin VB.Form frmUserinfo 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User info"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
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
   ScaleHeight     =   3135
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User's Name"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmUserinfo.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year of birth"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   315
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   840
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Index           =   3
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Index           =   4
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   330
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   840
      X2              =   5160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   3
      X1              =   840
      X2              =   5160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   840
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmUserinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub Form_Activate()
'SetWindowPos hWnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

