VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chatting with <name>"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrIsMeFocused 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   960
   End
   Begin MSComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   6330
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5424
            Picture         =   "frmChat.frx":0442
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      Picture         =   "frmChat.frx":059C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txtText 
      Height          =   645
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmChat.frx":06E6
   End
   Begin VB.Timer tmrIdle 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   360
      Top             =   960
   End
   Begin RichTextLib.RichTextBox txtReceived 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Mark the text you want to copy to the clipboard."
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0763
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3960
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":07E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0E16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbBar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1429
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Block User"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Dock Settings"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgInfo 
      Height          =   240
      Left            =   480
      Picture         =   "frmChat.frx":1130
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIdle 
      Height          =   240
      Left            =   240
      Picture         =   "frmChat.frx":127A
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgTyping 
      Height          =   240
      Left            =   0
      Picture         =   "frmChat.frx":13C4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send a file..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SentType As Boolean
Public Focused As Boolean
Public AppBar As New TAppBar

Private Sub cmdSend_Click()
On Error Resume Next
If GetPiece2(txtText.Text, " ", 1) = "" Then
    txtText.Text = ""
    Exit Sub
End If
txtText.Text = Replace(txtText.Text, vbCrLf, vbCr)
frmMain.SendData "outgoingmessage;" & GetPiece(Tag, vbLf, 2) & vbLf & txtText.Text
PutData txtReceived, "2" & IIf(frmMain.myName <> "", frmMain.myName, frmMain.myEmail) & "> " & txtText.Text, False
'txtReceived.OLEObjects.Add , , imgIdle.Picture
txtText.Text = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
'AppBar.Extends Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtReceived.Top = 120 + tbBar.Height
txtReceived.Left = 120
txtReceived.Width = ScaleWidth - 240
txtReceived.Height = ScaleHeight - 360 - txtText.Height - sbBar.Height - tbBar.Height
txtText.Left = 120
txtText.Top = txtReceived.Top + txtReceived.Height + 120
txtText.Width = txtReceived.Width - cmdSend.Width - 120
cmdSend.Left = txtText.Width + 240
cmdSend.Top = txtText.Top
End Sub


Function AddServerChat(Text As String)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = "*** " & Text
rt.SelStart = 0
rt.SelLength = Len("***")
rt.SelColor = vbRed
rt.SelBold = True
rt.SelStart = 0
rt.SelLength = 0
DoEvents
txtReceived.SelStart = Len(txtReceived.Text)
DoEvents
txtReceived.SelRTF = rt.TextRTF & vbCrLf


DoEvents
txtReceived.SelStart = Len(txtReceived)

rt.TextRTF = ""
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.SendData "closedchat;" & GetPiece(Tag, vbLf, 2)
AppBar.Detach
End Sub

Private Sub mnuClose_Click()
On Error Resume Next
Unload Me
End Sub


Private Sub mnuSendFile_Click()
On Error Resume Next
    frmMain.CreateNewFileSend GetPiece(Tag, vbLf, 2)
End Sub

Private Sub tbBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
    Case 1
        mnuSendFile_Click
    Case 3
        mnuBlockUser_click
    Case 5
        frmAppBar.Show
End Select
End Sub

Sub mnuBlockUser_click()
On Error Resume Next
Dim iUser As String
Dim rYN
If frmMain.wsock.State <> sckConnected Then Exit Sub
Err = 0
iUser = GetPiece(Tag, vbLf, 2)
If Err <> 0 Then Exit Sub
rYN = MsgBox("Are you absolutely sure you want to block '" & iUser & "'?", vbYesNo + vbQuestion, "Add user to block list.")
If rYN = vbNo Then Exit Sub
frmMain.SendData "blockuser;" & iUser
frmMain.SendData "whohasme;"

End Sub
Private Sub tmrIdle_Timer()
On Error Resume Next
SentType = False
frmMain.SendData "notyping;" & GetPiece(Tag, vbLf, 2)
tmrIdle.Enabled = False
End Sub

Private Sub tmrIsMeFocused_Timer()
On Error Resume Next
If GetForegroundWindow = hwnd Then
    tmrMessage.Enabled = False
    tmrIsMeFocused.Enabled = False
    Me.Caption = frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
End If
End Sub


Private Sub tmrMessage_Timer()
On Error Resume Next
If Me.Caption = frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) Then
    Me.Caption = "| - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
ElseIf Me.Caption = "| - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) Then
    Me.Caption = "\ - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
ElseIf Me.Caption = "\ - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) Then
    Me.Caption = "--  " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
ElseIf Me.Caption = "--  " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) Then
    Me.Caption = "/ - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
ElseIf Me.Caption = "/ - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) Then
    Me.Caption = "| - " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
End If
End Sub

Private Sub txtReceived_Change()
On Error Resume Next
Dim r As Long
If GetForegroundWindow <> hwnd Then
    r = FlashWindow(hwnd, 1)
End If
End Sub



Private Sub txtReceived_SelChange()
On Error Resume Next
If txtReceived.SelText <> "" Then
    Clipboard.Clear
    Clipboard.SetText txtReceived.SelText
End If
End Sub


Private Sub txtText_Change()
On Error Resume Next
If Len(txtText.Text) = 0 Then
    cmdSend.Enabled = False
    If SentType = True Then
        SentType = False
        frmMain.SendData "notyping;" & GetPiece(Tag, vbLf, 2)
        tmrIdle.Enabled = False
    End If
Else
    cmdSend.Enabled = True
    tmrIdle.Enabled = False
    tmrIdle.Enabled = True
    If SentType = False Then
        SentType = True
        frmMain.SendData "typing;" & GetPiece(Tag, vbLf, 2)
    End If
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
On Error Resume Next
        
        If KeyAscii = 11 Then
            txtText.SelText = Chr(Color)
            KeyAscii = 0
        ElseIf KeyAscii = 2 Then
            txtText.SelText = Chr(bold)
            KeyAscii = 0
        ElseIf KeyAscii = 21 Then
            txtText.SelText = Chr(underline)
            KeyAscii = 0
        ElseIf KeyAscii = 18 Then
            txtText.SelText = Chr(REVERSE)
            KeyAscii = 0
        ElseIf KeyAscii = 13 Then
'            KeyAscii = 0
        End If
End Sub
