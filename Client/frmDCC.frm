VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDCC 
   Caption         =   "Direct Connection Chat"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   735
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
   Icon            =   "frmDCC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgList 
      Left            =   2400
      Top             =   4800
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
            Picture         =   "frmDCC.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCC.frx":0464
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
         NumButtons      =   3
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
      EndProperty
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   5400
   End
   Begin VB.Timer tmrIsMeFocused 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   5400
   End
   Begin VB.CommandButton cmdSend 
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Picture         =   "frmDCC.frx":077E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtText 
      Height          =   645
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   4095
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   2160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrIdle 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   480
      Top             =   360
   End
   Begin MSComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
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
            Picture         =   "frmDCC.frx":08C8
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "12:07"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmDCC.frx":0A22
   End
   Begin RichTextLib.RichTextBox txtReceived 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Mark the text you want to copy to the clipboard."
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDCC.frx":0A9F
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
   Begin VB.Image imgTyping 
      Height          =   240
      Left            =   0
      Picture         =   "frmDCC.frx":0B1E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIdle 
      Height          =   240
      Left            =   240
      Picture         =   "frmDCC.frx":0C68
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgInfo 
      Height          =   240
      Left            =   480
      Picture         =   "frmDCC.frx":0DB2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSendFile 
         Caption         =   "&Send a file..."
      End
      Begin VB.Menu mnuSepdkjk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close this chat"
      End
   End
End
Attribute VB_Name = "frmDCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public SentType As Boolean

Sub CloseConnection()

End Sub

Sub mnuBlockUser_click()
On Error Resume Next
Dim iUser, rYN
If frmMain.wsock.State <> sckConnected Then Exit Sub
Err = 0
iUser = GetPiece(Tag, vbLf, 2)
If Err <> 0 Then Exit Sub
rYN = MsgBox("Are you absolutely sure you want to block '" & iUser & "'? To unblock", vbYesNo + vbQuestion, "Add user to block list.")
If rYN = vbNo Then Exit Sub
frmMain.SendData "blockuser;" & iUser
frmMain.SendData "whohasme;"

End Sub

Sub SendData(Text As String)
On Error Resume Next
If wsock.State <> sckConnected Then CloseConnection: Exit Sub
wsock.SendData Text & vbCrLf
DoEvents
End Sub


Private Sub cmdSend_Click()
On Error Resume Next
If GetPiece2(txtText.Text, " ", 1) = "" Then
    txtText.Text = ""
    Exit Sub
End If
txtText.Text = Replace(txtText.Text, vbCrLf, vbCr)
SendData "chat;" & txtText.Text
PutData txtReceived, "2" & IIf(frmMain.myName <> "", frmMain.myName, frmMain.myEmail) & "> " & txtText.Text, False
txtText.Text = ""
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
Function Confirm() As Boolean
Dim rYN
rYN = MsgBox("You have an incoming Direct Connection Chat Request from " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) & ". Would you like to accept it?", vbYesNo + vbQuestion, "Direct Connection")
If rYN = vbNo Then Confirm = False Else Confirm = True
End Function


Sub ChangeChatStatus(who As String, Typing As Boolean)
On Error Resume Next
        If Typing Then
            sbBar.Panels(1).Text = frmMain.GetNameFromEmail(who) & " is typing a message..."
            sbBar.Panels(1).Picture = imgTyping.Picture
            Exit Sub
        Else
            sbBar.Panels(1).Text = frmMain.GetNameFromEmail(who) & " is idle."
            sbBar.Panels(1).Picture = imgIdle.Picture
            Exit Sub
        End If
End Sub
Sub ChangeChatStatusBar(who As String, Text As String)
On Error Resume Next
sbBar.Panels(1).Text = Text
sbBar.Panels(1).Picture = imgInfo.Picture
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


Function AddChat(User As String, Text As String, Col As Boolean)
rt.SelBold = False
rt.SelColor = vbBlack

rt.SelStart = 0
rt.SelLength = 0
rt.SelRTF = User & "> " & Text
rt.SelStart = 0
rt.SelLength = Len(User)
If Col Then
    rt.SelColor = vbRed
Else
    rt.SelColor = vbBlue
End If
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

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuSendFile_Click()
On Error Resume Next
    frmMain.CreateNewFileSend GetPiece(Tag, vbLf, 2)
End Sub

Private Sub tbBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        mnuSendFile_Click
    Case 3
        mnuBlockUser_click
End Select
End Sub

Private Sub tmrIdle_Timer()
SentType = False
SendData "notyping;"
tmrIdle.Enabled = False
End Sub

Private Sub tmrIsMeFocused_Timer()
If GetForegroundWindow = hWnd Then
    tmrMessage.Enabled = False
    tmrIsMeFocused.Enabled = False
    Me.Caption = frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
End If
End Sub

Private Sub tmrMessage_Timer()
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
Dim r As Long
If GetForegroundWindow <> hWnd Then
    r = FlashWindow(hWnd, 1)
    'tmrIsMeFocused.Enabled = True
    'tmrMessage.Enabled = True
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
        SendData "notyping;"
        tmrIdle.Enabled = False
    End If
Else
    cmdSend.Enabled = True
    tmrIdle.Enabled = False
    tmrIdle.Enabled = True
    If SentType = False Then
        SentType = True
        SendData "typing;"
    End If
End If
End Sub


Private Sub wsock_Close()
wsock.Close
txtText.Text = ""
cmdSend.Enabled = False
txtText.Enabled = False
ChangeChatStatusBar frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)), "Direct Connection with " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) & " lost."
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
            KeyAscii = 0
        End If
End Sub

Private Sub wsock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
wsock.Close
wsock.Accept requestID
sbBar.Panels(1).Text = "Direct Connection Established."
End Sub

Private Sub wsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim lstItem As ListItem
Dim vtData As String
Dim curPOS As Single
Dim MESSAGE As String
Dim COMM As String
Dim rUser As String
Dim rTemp, iTemp, rTempV, rTmp
Dim rMsg As String
Static rApp
Static iCount
Static rREC
wsock.GetData vtData
vtData = vtData
wsock.Tag = wsock.Tag & vtData
GoTo CheckMSG
cmd:
    Select Case COMM
        Case "chat"
            PutData txtReceived, "4" & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) & "> " & MESSAGE, True
        Case "typing"
            ChangeChatStatus frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)), True
        Case "notyping"
            ChangeChatStatus frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)), False
    End Select


curPOS = 0
CheckMSG:
    rTemp = wsock.Tag
    If rTemp = "" Then Exit Sub
    
    Do
        iCount = iCount + 1
        
        iTemp = Mid(rTemp, iCount, 1)
        
        rTempV = wsock.Tag
        If Mid(rTemp, iCount, 2) = vbCrLf Then
                wsock.Tag = Right(rTemp, Len(rTemp) - (iCount + 1))
                iCount = 0
                COMM = ""
                MESSAGE = ""
                curPOS = 0
                Do
                    curPOS = curPOS + 1
                    rTmp = Left(rREC, curPOS)
                    rTmp = Right(rTmp, 1)
                    If rTmp = ";" Then Exit Do
                    COMM = COMM & rTmp
                Loop Until curPOS >= Len(rREC)
                COMM = LCase$(COMM)
                Do
                    If curPOS = Len(rREC) Then Exit Do
                    curPOS = curPOS + 1
                    rTmp = Left(rREC, curPOS)
                    rTmp = Right(rTmp, 1)
                    MESSAGE = MESSAGE & rTmp
                Loop Until curPOS >= Len(rREC)
                rTemp = ""
                rREC = ""
                GoTo cmd
            Else
                rREC = rREC & iTemp
        End If
        
Loop Until wsock.Tag = "" Or iCount >= Len(wsock.Tag)

End Sub

