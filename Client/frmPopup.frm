VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2205
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
   ScaleHeight     =   1575
   ScaleWidth      =   2205
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1200
      Top             =   0
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer tmrHeight 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   240
      Top             =   0
   End
   Begin VB.PictureBox picBack 
      Align           =   1  'Align Top
      Height          =   1575
      Left            =   0
      Picture         =   "frmPopup.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   0
      Width           =   2205
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "is online."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub InitPos()
Dim WindowRect As RECT
SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
Me.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - Me.Height
Me.Left = WindowRect.Right * Screen.TwipsPerPixelX - Me.Width
End Sub


Private Sub Form_Load()
Me.Height = 120
tmrHeight.Enabled = True
End Sub

Private Sub Label1_Click()
picBack_Click
End Sub

Private Sub lblUser_Click()
picBack_Click
End Sub

Private Sub picBack_Click()
On Error Resume Next
Dim frm As Form

Hide
DoEvents
If GetPiece(Tag, vbLf, 1) = "popup" Then
    If Not frmMain.ExistsChat(GetPiece(Tag, vbLf, 2)) Then
        frmMain.CreateNewChat GetPiece(Tag, vbLf, 2), False
    End If
ElseIf GetPiece(Tag, vbLf, 1) = "email" Then
Shell GetEmailProgram
ElseIf GetPiece(Tag, vbLf, 1) = "chatpopup" Then
    For Each frm In Forms
        If frm.Tag = "chat" & vbLf & GetPiece(Tag, vbLf, 2) Then
            frm.Show
            frm.WindowState = vbNormal
            Exit Sub
        End If
    Next
End If
Unload Me
End Sub

Private Sub tmrHeight_Timer()
Me.Height = Me.Height + 200
InitPos

If Me.Height >= 1665 Then
    Me.Height = 1665
    tmrHeight.Enabled = False
    tmrHide.Enabled = True
    InitPos
End If
End Sub


Private Sub tmrHide_Timer()
tmrHide.Enabled = False
tmrUnload = True
End Sub

Private Sub tmrUnload_Timer()
Me.Height = Me.Height - 200
InitPos

If Me.Height <= 120 Then
    Me.Height = 120
    Unload Me
End If
End Sub


