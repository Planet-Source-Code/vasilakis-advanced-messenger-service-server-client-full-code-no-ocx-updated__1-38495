VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3735
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
   ScaleHeight     =   975
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnim 
      Interval        =   150
      Left            =   3000
      Top             =   120
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnect.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnect.frx":0C54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      MouseIcon       =   "frmConnect.frx":18A8
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
   Begin VB.Image imgAnim 
      Height          =   480
      Left            =   240
      Picture         =   "frmConnect.frx":1BB2
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public mCounter
Private Sub Form_Load()
On Error Resume Next
Me.Left = (Screen.Width / 2) - (Width / 2)
Me.Top = (Screen.Height / 2) - (Height / 2)
SetWindowPos hwnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
mCounter = 1
End Sub

Private Sub lblCancel_Click()
Hide
frmMain.Enabled = True
frmMain.wsock.Close
End Sub

Private Sub tmrAnim_Timer()
On Error Resume Next

    mCounter = mCounter + 1
    imgAnim.Picture = imgList.ListImages(mCounter).Picture
    modIcon frmMain, frmMain.IconObject.Handle, frmMain.imgList.ListImages(mCounter).Picture, "vasilakis Messenger! - Connecting..."
    If mCounter > imgList.ListImages.Count - 1 Then mCounter = 0

End Sub


