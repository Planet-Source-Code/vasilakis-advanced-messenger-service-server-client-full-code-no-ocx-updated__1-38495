VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWhoHas 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
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
   ScaleHeight     =   3735
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnblock 
      Caption         =   "UnBlock User"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBlock 
      Caption         =   "Block User"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   2415
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   14941183
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Email"
         Object.Width           =   3755
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      Picture         =   "frmWhoHas.frx":0000
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoHas.frx":263F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoHas.frx":2BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoHas.frx":3173
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoHas.frx":370D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmWhoHas.frx":3CA7
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The following users has added you to their list and can see your status."
      Height          =   435
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   4065
   End
End
Attribute VB_Name = "frmWhoHas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlock_Click()
On Error Resume Next
If frmMain.wsock.State <> sckConnected Then Exit Sub
Err = 0
iUser = lstUsers.SelectedItem.Text
If Err <> 0 Then Exit Sub
If lstUsers.SelectedItem.SmallIcon = 4 Then Exit Sub
rYN = MsgBox("Are you absolutely sure you want to block '" & iUser & "'?", vbYesNo + vbQuestion, "Add user to block list.")
If rYN = vbNo Then Exit Sub
frmMain.SendData "blockuser;" & iUser
frmMain.SendData "whohasme;"
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub


Private Sub cmdUnblock_Click()
On Error Resume Next
If frmMain.wsock.State <> sckConnected Then Exit Sub
Err = 0
iUser = lstUsers.SelectedItem.Text
If Err <> 0 Then Exit Sub
If lstUsers.SelectedItem.SmallIcon <> 4 Then Exit Sub
rYN = MsgBox("Are you absolutely sure you want to remove from your block list user '" & iUser & "'?", vbYesNo + vbQuestion, "Add user to block list.")
If rYN = vbNo Then Exit Sub
frmMain.SendData "unblockuser;" & iUser
frmMain.SendData "whohasme;"
End Sub

Private Sub Form_Activate()
'SetWindowPos hWnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

