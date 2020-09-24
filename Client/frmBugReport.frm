VERSION 5.00
Begin VB.Form frmBugReport 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bug Report"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBugReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReport 
      Height          =   1455
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name"
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
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* The name field is required."
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   2040
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "* Please fill in the Report, describing of what happened, when, what did you do, etc."
      Height          =   435
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   4680
      Width           =   4920
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* The email field is required in case we have to contact with you."
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   4650
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bug report"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Email Address"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1620
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To complete this report, please fill in the required fields and press the 'Submit' button."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "This form helps you easily submit your bug reports, which help us improve OmniCom Messenger."
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmBugReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
On Error Resume Next
rYN = MsgBox("Are you sure you want to cancel this Bug Report?", vbYesNo + vbQuestion, "Cancel")
If rYN = vbYes Then
    Unload Me
End If
End Sub

Private Sub cmdSubmit_Click()
On Error Resume Next
frmMain.WebURL "http://liveupdate.vasilakis.gr/bugreport/send.asp?projectid=OMESSENGER&name=" & txtName.Text & "&from=" & txtEmail.Text & "&description=" & Replace(txtReport.Text, vbCrLf, "<BR>") & "&major_ver=" & App.Major & "&minor_ver=" & App.Minor & "&rev_ver=" & App.Revision
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
txtName.Text = frmMain.myName
txtEmail.Text = frmMain.myEmail
End Sub
