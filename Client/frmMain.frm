VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D3D3C3&
   Caption         =   "vasilakis :messenger v1.0"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   3570
   Begin VB.Frame frameToolBar 
      BackColor       =   &H00D3D3C3&
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   3570
      Begin VB.CommandButton cmdToolBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3C3&
         Caption         =   "Add"
         Height          =   780
         Index           =   0
         Left            =   45
         Picture         =   "frmMain.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   855
      End
      Begin VB.CommandButton cmdToolBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3C3&
         Caption         =   "Remove"
         Height          =   780
         Index           =   1
         Left            =   915
         Picture         =   "frmMain.frx":1A86
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   855
      End
      Begin VB.CommandButton cmdToolBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3C3&
         Caption         =   "Who?"
         Height          =   780
         Index           =   2
         Left            =   1785
         Picture         =   "frmMain.frx":26CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   45
         Width           =   855
      End
      Begin VB.CommandButton cmdToolBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3C3&
         Caption         =   "Settings"
         Height          =   780
         Index           =   3
         Left            =   2655
         Picture         =   "frmMain.frx":330E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.Timer tmrCheckEmail 
      Interval        =   30000
      Left            =   1800
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock wsockEMAIL 
      Left            =   600
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   100
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3C3&
      FillColor       =   &H00CDCAB9&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   3540
      TabIndex        =   11
      Top             =   6330
      Width           =   3570
      Begin VB.PictureBox picBannerHolder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00D3D3C3&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   480
         ScaleHeight     =   615
         ScaleMode       =   0  'User
         ScaleWidth      =   2895
         TabIndex        =   12
         Top             =   120
         Width           =   2895
         Begin VB.Image AnimatedGIF 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   0
            MouseIcon       =   "frmMain.frx":3F52
            MousePointer    =   99  'Custom
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.Timer AnimationTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox def_BANNER 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00D3D3C3&
         BorderStyle     =   0  'None
         FillColor       =   &H00CDCAB9&
         Height          =   870
         Left            =   0
         MouseIcon       =   "frmMain.frx":425C
         MousePointer    =   99  'Custom
         ScaleHeight     =   870
         ScaleWidth      =   2940
         TabIndex        =   13
         ToolTipText     =   "ruruy"
         Top             =   0
         Visible         =   0   'False
         Width           =   2940
      End
   End
   Begin VB.Timer tmrBanner 
      Interval        =   60000
      Left            =   2280
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock wsockBanner 
      Left            =   120
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7290
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Timer tmrActiveConnection 
      Interval        =   1000
      Left            =   3240
      Top             =   5880
   End
   Begin VB.Timer tmrAnimOnline 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   5880
   End
   Begin VB.Frame frameUsers 
      BackColor       =   &H00E3FBFF&
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   3915
      Begin VB.PictureBox picConnecting 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E3FBFF&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   3855
         TabIndex        =   8
         Top             =   360
         Width           =   3855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please Wait"
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
            Left            =   960
            TabIndex        =   9
            Top             =   240
            Width           =   990
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   360
            Picture         =   "frmMain.frx":4566
            Top             =   120
            Width           =   480
         End
      End
      Begin MSComctlLib.ListView lstWait 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   14941183
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSWinsockLib.Winsock wsock 
         Left            =   840
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstUsers 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   14941183
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
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3175
         EndProperty
         Picture         =   "frmMain.frx":4E30
      End
      Begin VB.Image cmdStatus 
         Height          =   240
         Left            =   1800
         Picture         =   "frmMain.frx":746F
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgMyStatus 
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":75B9
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblMyStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My Status"
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
         Left            =   405
         MouseIcon       =   "frmMain.frx":7B43
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblWait 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3FBFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Wait List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   4
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdConnect 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3C3&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   375
   End
   Begin MSComctlLib.ImageList imgListToolBar 
      Left            =   720
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8AA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A349
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1320
      Top             =   5760
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
            Picture         =   "frmMain.frx":AF9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B537
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BAD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C06B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "w w w . v a s i l a k i s . c o m"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1200
      MouseIcon       =   "frmMain.frx":C605
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   5520
      Width           =   2310
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   4320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lnUp 
      BorderColor     =   &H00928C67&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   4320
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
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
      Left            =   2640
      TabIndex        =   6
      Top             =   5310
      Width           =   1095
   End
   Begin VB.Image imgStat 
      Height          =   480
      Left            =   3480
      Picture         =   "frmMain.frx":C90F
      Top             =   5160
      Width           =   480
   End
   Begin VB.Label lblEmail1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3720
      TabIndex        =   5
      Top             =   5475
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgTrayOffLine 
      Height          =   240
      Left            =   840
      Picture         =   "frmMain.frx":CC19
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgTrayOnLine 
      Height          =   240
      Left            =   600
      Picture         =   "frmMain.frx":CD63
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDisconnect 
      Height          =   330
      Left            =   120
      Picture         =   "frmMain.frx":D2ED
      Top             =   5280
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgConnect 
      Height          =   330
      Left            =   120
      Picture         =   "frmMain.frx":D907
      Top             =   5280
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgOnline 
      Height          =   480
      Left            =   1560
      Picture         =   "frmMain.frx":DF23
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgOffline 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMain.frx":E22D
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Connection"
      Begin VB.Menu mnuLogIn 
         Caption         =   "Log in"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log out"
      End
      Begin VB.Menu mnuSepPopup 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowMessenger 
         Caption         =   "Show Messenger"
      End
      Begin VB.Menu mnuSepPopup2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "My Status"
         Begin VB.Menu mnuStatusOnline 
            Caption         =   "&Online"
         End
         Begin VB.Menu mnuSep321 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStatusAway 
            Caption         =   "&Away"
         End
         Begin VB.Menu mnuStatusBusy 
            Caption         =   "Bu&sy"
         End
         Begin VB.Menu mnuStatusBRB 
            Caption         =   "&Be Right Back"
         End
         Begin VB.Menu mnuSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMySettings 
            Caption         =   "My &Settings..."
         End
      End
      Begin VB.Menu mnuSep1233 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPopUser 
      Caption         =   "mnuPopUser"
      Visible         =   0   'False
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send file(s)"
      End
      Begin VB.Menu mnuDCC 
         Caption         =   "Direct Connection Chat"
      End
      Begin VB.Menu mnuSep123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserInfo 
         Caption         =   "User's Info"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSep3242332 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBugReport 
         Caption         =   "Problem/Bug Report"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public URL As String
Public ConnectCMD As Boolean
Public myEmail As String
Public myName As String
Public myYear
Public mySex
Public myCountry
Public myCity

Public mCounter
Public ofSec

Public JustLoggedIn As Boolean
Public IconObject As Object

Public EndingApp As Boolean

Public bannerdata As String
Public bannercurSize
Public bannerSize

Dim RepeatTimes&
Dim RepeatCount&
Dim FrameCount&
Dim TotalFrames&


Public emailUsername As String
Public emailPassword As String
Public emailPOPServer As String
Public emailLastEmails As String
Public emailCheckMail As Boolean


Public strBannerServer As String
Sub CenterBanners()
On Error Resume Next

picBannerHolder.Left = (picDown.ScaleWidth / 2) - (picBannerHolder.Width / 2)
picBannerHolder.Top = (picDown.ScaleHeight / 2) - (picBannerHolder.Height / 2)

End Sub

Function CountAllUsers() As Integer
On Error Resume Next
Dim lstItem As ListItem
Dim i
CountAllUsers = 0
For i = 1 To lstUsers.ListItems.Count
    Err = 0
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.SmallIcon = 2 Or lstItem.SmallIcon = 4 Then
        CountAllUsers = CountAllUsers + 1
    End If
Next i
For i = 1 To lstUsers.ListItems.Count
    Err = 0
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.SmallIcon = 1 Then
        CountAllUsers = CountAllUsers + 1
    End If
Next i
End Function

Sub CreateNewEmailPopup(Messages As String, bytes As String)
On Error Resume Next
Popups = Popups + 1
ReDim Preserve PopupWindow(Popups)
With PopupWindow(Popups)
    .Tag = "email" & vbLf & who
    .lblUser.Caption = Messages & " new mail(s)."
    .Label1.Caption = "(" & bytes & " bytes)"
    .InitPos
    SetWindowPos .hWnd, conHwndTopmost, .Left / 15, .Top / 15, .Width / 15, .Height / 15, conSwpNoActivate Or conSwpShowWindow
End With
End Sub

Sub CycleBanner()
On Error Resume Next
            UnloadBanner
            wsockBanner.Close
            wsockBanner.Connect strBannerServer, 8990
End Sub


Sub UnloadBanner()
    On Error Resume Next
    AnimationTimer.Enabled = False
    Dim anim As Image
    For Each anim In AnimatedGIF
        If anim.Index > 0 Then
            Unload anim
        End If
    Next anim
    AnimatedGIF(0).Picture = def_BANNER.Picture
    AnimatedGIF(0).Visible = True
    AnimatedGIF(0).ZOrder 0
    def_BANNER.Visible = False
    picBannerHolder.Height = AnimatedGIF(0).Height
    picBannerHolder.Width = AnimatedGIF(0).Width
    URL = "http://www.vasilakis.com/messenger.asp"
    CenterBanners
End Sub

Private Sub AnimatedGIF_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then
    UnloadBanner
    tmrBanner.Enabled = False
    tmrBanner.Enabled = True
Else
    If URL <> "" Then
            WebURL URL
            tmrBanner_Timer
            tmrBanner.Enabled = False
            tmrBanner.Enabled = True
    End If
End If
End Sub

Private Sub AnimationTimer_Timer()
On Error Resume Next

    If FrameCount < TotalFrames Then
        FrameCount = FrameCount + 1
        AnimatedGIF(FrameCount).Visible = True
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    Else
        FrameCount = 0
        For i = 1 To AnimatedGIF.Count - 1
            AnimatedGIF(i).Visible = False
        Next i
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    End If
       
End Sub

Sub AddChatText(who As String, Text As String)
On Error Resume Next
Dim frm As Form
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then
        PutData frm.txtReceived, "4" & GetNameFromEmail(who) & "> " & Text, True
        frm.Caption = GetNameFromEmail(who)
        Exit Sub
    End If
Next
End Sub


Sub ChangeChatStatusBar(who As String, Text As String)
On Error Resume Next
Dim frm As Form
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then
        frm.sbBar.Panels(1).Text = Text
        frm.sbBar.Panels(1).Picture = frm.imgInfo.Picture
        Exit Sub
    End If
Next
End Sub


Sub ChangeChatStatus(who As String, Typing As Boolean)
On Error Resume Next
Dim frm As Form
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then
        If Typing Then
            frm.sbBar.Panels(1).Text = GetNameFromEmail(who) & " is typing a message..."
            frm.sbBar.Panels(1).Picture = frm.imgTyping.Picture
            Exit Sub
        Else
            frm.sbBar.Panels(1).Text = GetNameFromEmail(who) & " is idle."
            frm.sbBar.Panels(1).Picture = frm.imgIdle.Picture
            Exit Sub
        End If
    End If
Next

End Sub
Sub AddServerChatText(who As String, Text As String)
On Error Resume Next
Dim frm As Form
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then
        frm.AddServerChat Text
        Exit Sub
    End If
Next
End Sub
Sub AddToList(who As String, Online As Boolean)
On Error Resume Next
Dim lstItem As ListItem

RemoveFromList who
tmrAnimOnline.Enabled = False
mCounter = 0
ofSec = 0
If Online Then
    RemoveNoneCaption
    Set lstItem = lstUsers.ListItems.Add(2, , who, , 2)
    'SendData "getname;" & lstItem.Text
Else
    RemoveOfflineCaption
    Set lstItem = lstUsers.ListItems.Add(GetOffLineIndex + 1, , who, , 1)
    SendData "getname;" & lstItem.Text
'    Show
End If
lstUsers.Refresh
End Sub

Sub AddToWait(who As String)
On Error Resume Next
Dim lstItem As ListItem
RemoveFromList who
If lstWait.ListItems(1).Text = "None." Then
    lstWait.ListItems.Clear
End If
Set lstItem = lstWait.ListItems.Add(1, , who, , 3)
SendData "getname;" & lstItem.Text
lstWait.Refresh
End Sub
Sub ChangeDisplayName(FromWho As String, NName As String)
On Error Resume Next
Dim lstItem As ListItem
Dim i
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = FromWho Then
        lstItem.SubItems(1) = NName
        Exit For
    End If
Next i
For i = 1 To lstWait.ListItems.Count
    Set lstItem = lstWait.ListItems(i)
    If lstItem.Text = FromWho Then
        lstItem.SubItems(1) = NName
        Exit For
    End If
Next i
End Sub

Sub ChangeUserStatus(FromWho As String, status As Integer)
On Error Resume Next
Dim lstItem As ListItem
Dim i
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = FromWho Then
        If status = 0 Then
            lstItem.SmallIcon = 2
            SendData "getname;" & FromWho
        ElseIf status = 1 Then
            lstItem.SmallIcon = 4
            lstItem.SubItems(1) = lstItem.SubItems(1) & " (Away)"
        ElseIf status = 2 Then
            lstItem.SmallIcon = 4
            lstItem.SubItems(1) = lstItem.SubItems(1) & " (Busy)"
        ElseIf status = 3 Then
            lstItem.SmallIcon = 4
            lstItem.SubItems(1) = lstItem.SubItems(1) & " (Be Right Back)"
        End If
        lstUsers.Refresh
        Exit For
    End If
Next i
End Sub


Sub CloseConnection()
On Error Resume Next
lstUsers.ListItems.Clear
lstWait.ListItems.Clear
DoEvents
picConnecting.Visible = False
mnuLogIn.Enabled = True
mnuLogOut.Enabled = False
mnuStatus.Enabled = False
Dim lstItem As ListItem
lstUsers.ListItems.Clear
lstWait.ListItems.Clear
lblName.Caption = "Offline"
imgStat.Picture = imgOffline.Picture
cmdConnect.Picture = imgConnect.Picture
cmdConnect.ToolTipText = "Connect to host"
modIcon Me, IconObject.Handle, imgTrayOffLine.Picture, "vasilakis Messenger! - Offline"
cmdStatus.Enabled = False
lblMyStatus.Caption = "My Status"
imgMyStatus.Picture = imgList.ListImages(1).Picture
Set lstItem = lstUsers.ListItems.Add(, , "Not Connected.")
Set lstItem = lstWait.ListItems.Add(, , "Not Connected.")
sbBar.SimpleText = "Not Connected."
wsock.Close
'lstItem.SubItems(1) = "Not Connected."
'lstUsers.Enabled = False
End Sub

Sub CreateNewChat(who As String, minimized As Boolean)
On Error Resume Next
Dim rX As Integer, rY As Integer
ChatsIndex = ChatsIndex + 1
ReDim Preserve ChatForm(ChatsIndex)
ChatForm(ChatsIndex).Tag = "chat" & vbLf & who
ChatForm(ChatsIndex).Caption = GetNameFromEmail(who)
ChatForm(ChatsIndex).sbBar.Panels(1).Text = "Chatting with " & GetNameFromEmail(who) & "..."
ChatForm(ChatsIndex).Top = (Screen.Height / 2) - (ChatForm(ChatsIndex).Height / 2)
ChatForm(ChatsIndex).Left = (Screen.Width / 2) - (ChatForm(ChatsIndex).Width / 2)
If minimized Then
    ChatForm(ChatsIndex).WindowState = vbMinimized
End If
ChatForm(ChatsIndex).Visible = True
End Sub

Sub CreateNewDCChat(who As String, listen As Boolean, Optional host As String, Optional port As Double)
On Error Resume Next
Randomize Timer
Dim iPort As Double
Dim rX As Integer, rY As Integer
DCChatsIndex = DCChatsIndex + 1
ReDim Preserve DCChatForm(DCChatsIndex)
With DCChatForm(DCChatsIndex)
    .Tag = "dcchat" & vbLf & who
    .Caption = GetNameFromEmail(who) & " - Direct Connection"
    DoEvents
Err = 0
If listen = True Then
    .Show
    DoEvents
    .sbBar.Panels(1).Text = "Binding Port..."
    DoEvents
ReFindPort:
    iPort = Int(Rnd * 65000) + 1
    DoEvents
    Err = 0
    .wsock.Close
    .wsock.RemotePort = 0
    .wsock.LocalPort = iPort
    .wsock.listen
    DoEvents
    If Err <> 0 Then GoTo ReFindPort
    DoEvents
    .sbBar.Panels(1).Text = "Waiting for Connection..."
    DoEvents
    SendData "dcc;" & who & vbLf & iPort
Else
    If Not .Confirm Then
        Unload DCChatForm(DCChatsIndex)
        Exit Sub
    End If
    .Show
    DoEvents
    DoEvents
    .sbBar.Panels(1).Text = "Connecting..."
    DoEvents
    .wsock.Close
    .wsock.RemotePort = 0
    .wsock.LocalPort = 0
    .wsock.Connect host, port
    Do
    DoEvents
    Loop Until wsock.State = sckConnected Or wsock.State = sckError Or wsock.State = sckClosed
    Select Case wsock.State
        Case sckClosed, sckError
            MsgBox "Couldn't establish a direct connection to " & GetNameFromEmail(who) & ".", vbCritical, "Direct Connection"
            Unload DCChatForm(DCChatsIndex)
            Exit Sub
        Case sckConnected
            DoEvents
            .sbBar.Panels(1).Text = "Direct Connection Established."
    End Select
    DoEvents
End If
End With
End Sub
Sub CreateNewFileSend(who As String)
On Error Resume Next
Randomize Timer
Dim iPort As Double
Dim iFiles As String
Dim i
Dim rX As Integer, rY As Integer
FilesSendIndex = FilesSendIndex + 1
ReDim Preserve FileSend(FilesSendIndex)
With FileSend(FilesSendIndex)
    .Tag = "sendfile" & vbLf & who
    .Caption = "Sending File(s) to " & GetNameFromEmail(who)
    .Show
    'Set .IconObjectSend = .Icon
    'AddIcon FileSend(FilesSendIndex), .IconObjectSend.Handle, .IconObjectSend, "Waiting for connection..."
    
    DoEvents
Err = 0
    .lblStatus.Caption = "Binding Port..."
    DoEvents
ReFindPort:
    iPort = Int(Rnd * 65000) + 1
    DoEvents
    Err = 0
    .wsock.Close
    .wsock.RemotePort = 0
    .wsock.LocalPort = iPort
    .wsock.listen
    DoEvents
    If Err <> 0 Then GoTo ReFindPort
    DoEvents
    .lblStatus.Caption = "Choosing File(s)..."
    DoEvents
    .ChooseFiles
    If .lstFiles.ListCount = 0 Then
        Unload FileSend(FilesSendIndex)
        Exit Sub
    End If
    iFiles = ""
    For i = 1 To .lstFiles.ListCount
        iFiles = iFiles & .lstFiles.List(i) & "||"
    Next i
    .lblStatus.Caption = "Waiting for Connection..."
    
    SendData "sendfiles;" & who & vbLf & iPort & iFiles

End With

End Sub

Sub CreateNewFileReceive(who As String, host As String, port As Double)
On Error Resume Next
Randomize Timer
Dim iPort As Double
FilesIndex = FilesIndex + 1
ReDim Preserve FileTransfer(FilesIndex)
With FileTransfer(FilesIndex)
    .Tag = "receivefile" & vbLf & who
    .Caption = "Getting File(s) from " & GetNameFromEmail(who) & "."
    Set .IconObjectReceive = .Icon
    AddIcon FileTransfer(FilesIndex), .IconObjectReceive.Handle, .IconObjectReceive, "File Receipient."

    If Not .Confirm Then Unload FileTransfer(FilesIndex): Exit Sub
    .Show
    DoEvents
    .wsock.RemotePort = 0
    .wsock.LocalPort = 0
    .wsock.Connect host, port
    Do
        DoEvents
    Loop Until .wsock.State = sckConnected Or .wsock.State = sckError Or .wsock.State = sckClosed
    If .wsock.State = sckConnected Then
        .SendData "start" & vbCrLf
    Else
        Unload FileTransfer(FilesIndex)
    End If
    
    DoEvents

End With

End Sub


Sub CreateNewPopupWindow(who As String)
On Error Resume Next
Popups = Popups + 1
ReDim Preserve PopupWindow(Popups)
With PopupWindow(Popups)
    .Tag = "popup" & vbLf & who
    .lblUser.Caption = GetNameFromEmail(who)
    .InitPos
    SetWindowPos .hWnd, conHwndTopmost, .Left / 15, .Top / 15, .Width / 15, .Height / 15, conSwpNoActivate Or conSwpShowWindow
End With
End Sub

Sub CreateNewChatPopupWindow(who As String)
On Error Resume Next
Popups = Popups + 1
ReDim Preserve PopupWindow(Popups)
With PopupWindow(Popups)
    .Tag = "chatpopup" & vbLf & who
    .lblUser.Caption = GetNameFromEmail(who)
    .Label1.Caption = "has sent you a message."
    .InitPos
    SetWindowPos .hWnd, conHwndTopmost, .Left / 15, .Top / 15, .Width / 15, .Height / 15, conSwpNoActivate Or conSwpShowWindow
End With
End Sub

Sub EndApp()
On Error Resume Next
EndingApp = True
wsock.Close
CloseConnection
DoEvents
Hide
delIcon IconObject.Handle
DoEvents
Dim frm As Form
For Each frm In Forms
    If frm.Name <> Name Then
        Unload frm
    End If
Next
DoEvents
Unload Me
End Sub

Function GetOffLineIndex() As Integer
On Error Resume Next
Dim i
Dim lstItem As ListItem
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = "Not Online" Then
        GetOffLineIndex = lstItem.Index
        Exit Function
    End If
Next i
GetOffLineIndex = lstUsers.ListItems.Count
End Function

Function GetOnLineIndex() As Integer
On Error Resume Next
Dim i
Dim lstItem As ListItem
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = "Online" Then
        GetOnLineIndex = lstItem.Index
        Exit Function
    End If
Next i
GetOnLineIndex = lstUsers.ListItems.Count
End Function

Sub InitConnection()
On Error Resume Next
CloseConnection
wsock.Close
DoEvents
lstUsers.ListItems.Clear
lstWait.ListItems.Clear
Unload frmLogin
frmLogin.txtUSER.Text = GetSetting("vasilakis Messenger", "Login", "Email", "user@domain.com")

frmLogin.chkSavePassword.Value = GetSetting("vasilakis Messenger", "Login", "SavePass", "0")
If frmLogin.chkSavePassword.Value = 1 And ConnectCMD = False Then
    frmLogin.txtPASS.Text = GetSetting("vasilakis Messenger", "Login", "Password", "")
    'Visible = False
Else
    frmLogin.Show 1
End If
Unload frmSplash
ConnectCMD = False
myEmail = LCase$(frmLogin.txtUSER.Text)
'lblEmail.Caption = myEmail
If frmLogin.Cancel = True Then Exit Sub
mnuLogIn.Enabled = False
mnuLogOut.Enabled = False
mnuStatus.Enabled = False

picConnecting.Visible = True
Load frmConnect
frmConnect.Show
DoEvents
frmConnect.Refresh
DoEvents
wsock.Close
DoEvents
JustLoggedIn = True
lstUsers.ListItems.Clear
lstWait.ListItems.Clear
sbBar.SimpleText = "Connecting..."
wsock.Connect frmLogin.cboServer.Text, "9811"
Do
    DoEvents
Loop Until wsock.State = sckConnected Or wsock.State = sckClosed Or wsock.State = sckError
Unload frmConnect
DoEvents
If wsock.State <> sckConnected Then
    ConnectCMD = True
    CloseConnection
End If
DoEvents
'Enabled = True
'CreateNewFileSend "test"
'NewChat "Vasilakis"
End Sub

Function IsInList(who As String) As Boolean
On Error Resume Next
Dim lstItem As ListItem
Dim i
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = who Then
        IsInList = True
        Exit Function
    End If
Next i
For i = 1 To lstWait.ListItems.Count
    Set lstItem = lstWait.ListItems(i)
    If lstItem.Text = who Then
        IsInList = True
        Exit Function
    End If
Next i
End Function

Sub NewAuthorization(who As String)
On Error Resume Next
AuthsIndex = AuthsIndex + 1
ReDim Preserve AuthorizeUser(AuthsIndex)
AuthorizeUser(AuthsIndex).Tag = "auth" & vbLf & who
AuthorizeUser(AuthsIndex).lblStat.Caption = "User '" & who & "' has added you to his/her list."
If IsInList(who) Then
    AuthorizeUser(AuthsIndex).chkAdd.Enabled = False
End If
AuthorizeUser(AuthsIndex).Show
End Sub


Function ExistsChat(who As String) As Boolean
On Error Resume Next
Dim frm As Form
ExistsChat = False
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then ExistsChat = True: Exit Function
Next
End Function

Sub ShowChat(who As String)
On Error Resume Next
Dim frm As Form
For Each frm In Forms
    If frm.Tag = "chat" & vbLf & who Then frm.Show: Exit Sub
Next
End Sub


Function GetNameFromEmail(Email As String) As String
On Error Resume Next
Dim lstItem As ListItem
Dim i
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = Email Then
        If lstItem.SubItems(1) <> "" Then
            GetNameFromEmail = lstItem.SubItems(1)
        Else
            GetNameFromEmail = Email
        End If
        Exit Function
    End If
    DoEvents
Next i
For i = 1 To lstWait.ListItems.Count
    Set lstItem = lstWait.ListItems(i)
    If lstItem.Text = Email Then
        If lstItem.SubItems(1) <> "" Then
            GetNameFromEmail = lstItem.SubItems(1)
        Else
            GetNameFromEmail = Email
        End If
        Exit Function
    End If
    DoEvents
Next i
GetNameFromEmail = Email
End Function

Sub GetNamesList()
On Error Resume Next
Dim lstItem As ListItem
Dim i
For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    SendData "getname;" & lstItem.Text
    DoEvents
Next i
End Sub

Sub NewChat(who As String)
On Error Resume Next
If ExistsChat(who) Then ShowChat who: Exit Sub
If LCase$(who) = LCase$(myEmail) Then Exit Sub
CreateNewChat who, True
End Sub

Sub RemoveFromList(who As String)
On Error Resume Next
Dim lstItem As ListItem
Dim i
    If lstUsers.ListItems(1).Text = "None." Then
        lstUsers.ListItems.Remove 1
    End If

For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = who Then
        lstUsers.ListItems.Remove i
        Exit For
    End If
Next i
For i = 1 To lstWait.ListItems.Count
    Set lstItem = lstWait.ListItems(i)
    If lstItem.Text = who Then
        lstWait.ListItems.Remove i
        Exit For
    End If
Next i
If CountOnline = 0 And lstUsers.ListItems(2).Text <> "None." Then
    Set lstItem = lstUsers.ListItems.Add(2, , "None.")
End If
If CountOffline = 0 And lstUsers.ListItems(GetOffLineIndex + 1).Text <> "None." Then
    Set lstItem = lstUsers.ListItems.Add(GetOffLineIndex + 1, , "None.")
End If
If CountWait = 0 Then
    Set lstItem = lstWait.ListItems.Add(1, , "None.")
End If
End Sub

Sub RemoveNoneCaption()
On Error Resume Next
Dim lstItem As ListItem
Dim i

For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = "None." Then
        lstUsers.ListItems.Remove i
        Exit For
    End If
    If lstItem.Text = "Not Online" Then Exit For
Next i
End Sub

Sub RemoveOfflineCaption()
On Error Resume Next
Dim lstItem As ListItem
Dim i

For i = 1 To lstUsers.ListItems.Count
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.Text = "None." And lstItem.Index > GetOffLineIndex Then
        lstUsers.ListItems.Remove i
        Exit For
    End If
Next i

End Sub


Function CountOnline() As Integer
On Error Resume Next
Dim lstItem As ListItem
Dim i
CountOnline = 0
For i = 1 To lstUsers.ListItems.Count
    Err = 0
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.SmallIcon = 2 Or lstItem.SmallIcon = 4 Then
        CountOnline = CountOnline + 1
    End If
Next i
End Function

Function CountWait() As Integer
On Error Resume Next
Dim lstItem As ListItem
Dim i
CountWait = 0
For i = 0 To lstWait.ListItems.Count
    Err = 0
    Set lstItem = lstWait.ListItems(i)
    If Err = 0 Then
        CountWait = CountWait + 1
    End If
Next i
End Function



Function CountOffline() As Integer
On Error Resume Next
Dim lstItem As ListItem
Dim i
CountOffline = 0
For i = 1 To lstUsers.ListItems.Count
    Err = 0
    Set lstItem = lstUsers.ListItems(i)
    If lstItem.SmallIcon = 1 Then
        CountOffline = CountOffline + 1
    End If
Next i
End Function




Sub TogleStatusMenu(rStatus)
On Error Resume Next
mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusBusy.Checked = False
mnuStatusBRB.Checked = False

If rStatus = 0 Then
    mnuStatusOnline.Checked = True
    modIcon Me, IconObject.Handle, imgTrayOnLine.Picture, "vasilakis Messenger! - Online"
    sbBar.SimpleText = "Online."
ElseIf rStatus = 1 Then
    mnuStatusAway.Checked = True
    modIcon Me, IconObject.Handle, imgList.ListImages(4).Picture, "vasilakis Messenger! - (Away)"
    sbBar.SimpleText = "Away..."
ElseIf rStatus = 2 Then
    mnuStatusBusy.Checked = True
    modIcon Me, IconObject.Handle, imgList.ListImages(4).Picture, "vasilakis Messenger! - (Busy)"
    sbBar.SimpleText = "Busy..."
ElseIf rStatus = 3 Then
    mnuStatusBRB.Checked = True
    modIcon Me, IconObject.Handle, imgList.ListImages(4).Picture, "vasilakis Messenger! - (Be Right Back)"
    sbBar.SimpleText = "Be Right Back..."
End If

End Sub

Private Sub cmdConnect_Click()
On Error Resume Next
ConnectCMD = True
If cmdConnect.Picture = imgConnect.Picture Then
    InitConnection
Else
    wsock_Close
End If
End Sub



Private Sub cmdStatus_Click()
On Error Resume Next
PopupMenu mnuStatus
End Sub

Private Sub cmdToolBar_Click(Index As Integer)
On Error Resume Next
Dim lstItem As ListItem
Dim rUser As String
Dim rYN

Select Case Index
    Case 0
        If wsock.State <> sckConnected Then Exit Sub
        Unload frmAddUser
        frmAddUser.Show 1
        If frmAddUser.Cancel = True Then Exit Sub
        If LCase$(frmAddUser.txtUSER.Text) <> LCase$(myEmail) Then
            SendData "adduser;" & LCase$(frmAddUser.txtUSER.Text)
        Else
            MsgBox "You can't add yourself to your list.", vbCritical, "Error"
        End If
    Case 1
        If wsock.State <> sckConnected Then Exit Sub
        Set lstItem = lstUsers.SelectedItem
        If lstItem.Text = "Online" Then Exit Sub
        If lstItem.Text = "Not Connected." Then Exit Sub
        If lstItem.Text = "Not Online" Then Exit Sub
        If lstItem.Text = "None." Then Exit Sub
        Err = 0
        rUser = lstItem.Text
        If Err <> 0 Then Exit Sub
        rYN = MsgBox("Are you sure you want to remove user '" & GetNameFromEmail(rUser) & "' from you list, and never see his/her online status again?", vbYesNo + vbInformation, "Remove User")
        If rYN = vbYes Then
            SendData "removefromallowlist;" & rUser
        End If
    Case 2
        If wsock.State <> sckConnected Then Exit Sub
        SendData "whohasme;"
    Case 3
        If wsock.State = sckConnected Then
            frmSettings.txtName.Text = myName
            frmSettings.cboYear.Text = myYear
            frmSettings.cboSex.Text = mySex
            frmSettings.txtCountry.Text = myCountry
            frmSettings.txtCity.Text = myCity
            frmSettings.frameSettings(1).Visible = False
            frmSettings.frameSettings(0).Visible = True
            frmSettings.cmdFrame(0).FontBold = True
        Else
            frmSettings.cmdFrame(1).FontBold = True
            frmSettings.cmdFrame(0).Enabled = False
            frmSettings.frameSettings(0).Visible = False
            frmSettings.frameSettings(1).Visible = True
        End If
        
        frmSettings.Show
        frmSettings.txtName.SetFocus
End Select
End Sub



Public Function WebURL(ByVal URL As String) As Long
On Error Resume Next
    WebURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub def_BANNER_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then
    UnloadBanner
Else
    If URL <> "" Then
         WebURL URL
    End If
End If
End Sub


Private Sub Form_Load()
On Error Resume Next
strBannerServer = "localhost"
rAutoCheckMail = GetSetting("vasilakis Messenger", "Settings", "AutoCheckMail", "0")
If rAutoCheckMail = "1" Then
    emailCheckMail = True
    tmrCheckEmail.Enabled = True
    emailUsername = GetSetting("vasilakis Messenger", "Settings", "emailUsername")
    emailPassword = DecryptPassword(0.06, GetSetting("vasilakis Messenger", "Settings", "emailPassword"))
    emailPOPServer = GetSetting("vasilakis Messenger", "Settings", "emailPOPServer")
Else
    emailCheckMail = False
    tmrCheckEmail.Enabled = False
End If

tmrCheckEmail_Timer
URL = "http://liveupdate.vasilakis.com/omni_msg"
UnloadBanner
Me.Left = (Screen.Width / 2) - (Width / 2)
Me.Top = (Screen.Height / 2) - (Height / 2)
If App.PrevInstance Then End
Form_Resize
DoEvents
    strBold = Chr(bold)
    strUnderline = Chr(underline)
    strColor = Chr(Color)
    strReverse = Chr(REVERSE)
    strAction = Chr(ACTION)
Set IconObject = imgTrayOnLine.Picture
AddIcon Me, IconObject.Handle, IconObject, "vasilakis Messenger!"
DoEvents
CloseConnection
rDialUp = GetSetting("vasilakis Messenger", "Login", "DialUp", "0")
frmSplash.Hide
Unload frmSplash
DoEvents
If rDialUp = "1" Then
    tmrActiveConnection.Enabled = True
Else
    tmrActiveConnection.Enabled = False
    InitConnection
End If
'frmBar.Show
'picConnecting.Visible = True
End Sub






Sub CheckMail()
On Error Resume Next
If emailUsername = "" Then
    Exit Sub
End If
If emailPassword = "" Then
    Exit Sub
End If
If emailPOPServer = "" Then
    Exit Sub
End If
On Error Resume Next
Err = 0
wsockEMAIL.Close
wsockEMAIL.Tag = "NONE"
wsockEMAIL.RemotePort = 110
wsockEMAIL.LocalPort = 0
wsockEMAIL.Connect emailPOPServer
End Sub
Function GetParms(from As String, delim As String) As Long
On Error Resume Next
If Len(from) = 0 Then GetParms = 0: Exit Function
GetParms = 0
For i = 1 To Len(from)
    If Mid(from, i, 1) = delim Then GetParms = GetParms + 1
Next i
End Function


Sub LoadAniGif(xImgArray)
On Error Resume Next
picBannerHolder.Visible = False
    Dim F1, F2
    Dim AnimatedGIFs() As String
    Dim imgHeader As String
    Static buf$, picbuf$
    Dim fileHeader As String
    Dim imgCount
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd
    GifEnd = Chr(0) & "!ù"
    
    AnimationTimer.Enabled = False
    For i = 1 To xImgArray.Count - 1
        Unload xImgArray(i)
    Next i
    
    F1 = FreeFile
On Error GoTo badFile:
    buf = bannerdata
    
    i = 1
    imgCount = 0
    
    j = (InStr(1, buf, GifEnd) + Len(GifEnd)) - 2
    fileHeader = Left(buf, j)
    i = j + 2
    
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * CLng(256))
    Else
        RepeatTimes = 0
    End If


    Do
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + Len(GifEnd)
        If j > Len(GifEnd) Then
            F2 = FreeFile
            Open "tmp.gif" For Binary As F2
                picbuf = String(Len(fileHeader) + j - i, Chr(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #F2, 1, picbuf
                imgHeader = Left(Mid(buf, i - 1, j - i), 16)
            Close F2
            
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
                Load xImgArray(imgCount - 1)
                xImgArray(imgCount - 1).ZOrder 0
                xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
                xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
            End If
            xImgArray(imgCount - 1).Tag = TimeWait
            xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
            Kill ("tmp.gif")
            
            i = j '+ 1
        End If
        DoEvents
    Loop Until j = Len(GifEnd)
    
    If i < Len(buf) Then
        F2 = FreeFile
        Open "tmp.gif" For Binary As F2
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #F2, 1, picbuf
            imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close F2

        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
            Load xImgArray(imgCount - 1)
            xImgArray(imgCount - 1).ZOrder 0
            xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
            xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
        End If
        xImgArray(imgCount - 1).Tag = TimeWait
        xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
        Kill ("tmp.gif")
    End If
    
    FrameCount = 0
    TotalFrames = xImgArray.Count - 1
    picBannerHolder.Height = AnimatedGIF(0).Height
    picBannerHolder.Width = AnimatedGIF(0).Width
    CenterBanners
    picBannerHolder.Visible = True
On Error GoTo badTime
    AnimationTimer.Interval = CInt(xImgArray(0).Tag)
badTime:
    AnimationTimer.Enabled = True
Exit Sub
badFile:
    UnloadBanner
End Sub



Private Sub tmrCheckEmail_Timer()
On Error Resume Next
If emailCheckMail = True Then CheckMail
End Sub

Private Sub wsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wsock_Close
End Sub

Private Sub wsockEMAIL_DataArrival(ByVal bytesTotal As Long)
Dim vtData As String
Dim COMM As String
On Error Resume Next
Err = 0
wsockEMAIL.GetData vtData
If Err <> 0 Then Exit Sub
Do
   COMM = ""
   Do
    curPOS = curPOS + 1
    rTmp = Mid(vtData, curPOS, 1)
    If rTmp = Chr$(13) Then Exit Do
    COMM = COMM & rTmp
   Loop Until curPOS >= bytesTotal
   Err = 0
   If LCase$(GetPiece(COMM, " ", 1)) = "+ok" Then
        Select Case LCase$(wsockEMAIL.Tag)
            Case "none"
                wsockEMAIL.Tag = "USER"
                Err = 0
                wsockEMAIL.SendData "USER " & emailUsername & vbCrLf
            Case "user"
                wsockEMAIL.Tag = "PASS"
                Err = 0
                wsockEMAIL.SendData "PASS " & emailPassword & vbCrLf
            Case "pass"
                Err = 0
'                iTMP = 0
'                wsockEMAIL.Tag = "LIST"
'                wsockEMAIL.SendData "LIST" & vbCrLf
                wsockEMAIL.Tag = "STAT"
                wsockEMAIL.SendData "STAT" & vbCrLf
            Case "stat"
            Dim iTMP As String
                iTMP = Val(GetPiece(COMM, " ", 2))
                wsockEMAIL.Tag = "none"
                If iTMP > 0 Then
                    If emailLastEmails <> iTMP & GetPiece(COMM, " ", 3) Then
                        emailLastEmails = iTMP & GetPiece(COMM, " ", 3)
                        CreateNewEmailPopup iTMP, GetPiece(COMM, " ", 3)
                    End If
                Else
                End If
                wsockEMAIL.Tag = ""
                Err = 0
                wsockEMAIL.SendData "QUIT" & vbCrLf
        End Select
   ElseIf LCase$(GetPiece(COMM, " ", 1)) = "." Then
        Select Case LCase$(wsockEMAIL.Tag)
            Case "list"
                wsockEMAIL.Tag = "none"
                If iTMP > 0 Then
                    sb.SimpleText = iTMP & " Email(s) waiting."
                    Err = 0
                End If
                wsockEMAIL.Tag = ""
                wsockEMAIL.SendData "QUIT" & vbCrLf
        End Select
   ElseIf LCase$(GetPiece(COMM, " ", 1)) = "-err" Then
        Select Case LCase$(wsockEMAIL.Tag)
            Case "none"
                tmrCheckEmail.Enabled = False
                emailPOPServer = ""
                emailUsername = ""
                emailPassword = ""
                sb.SimpleText = "Server Error!"
                wsockEMAIL.Tag = ""
                Err = 0
                MsgBox "Cannot check email. Unknown server error.", vbCritical, "Mail check error"
                wsockEMAIL.SendData "QUIT" & vbCrLf
                MainEmailOptions
                Exit Sub
            Case "user"
                tmrCheckEmail.Enabled = False
                emailUsername = ""
                emailPassword = ""
                sb.SimpleText = "Username Invalid!"
                wsockEMAIL.Tag = ""
                Err = 0
                wsockEMAIL.SendData "QUIT" & vbCrLf
                MsgBox "Cannot check email. Username is invalid.", vbCritical, "Mail check error"
                MainEmailOptions
                Exit Sub
            Case "pass"
                tmrCheckEmail.Enabled = False
                emailPassword = ""
                wsockEMAIL.Tag = ""
                Err = 0
                wsockEMAIL.SendData "QUIT" & vbCrLf
                MsgBox "Cannot check email. Password is invalid.", vbCritical, "Mail check error"
                MainEmailOptions
                Exit Sub
        End Select
   Else
        Select Case LCase$(wsockEMAIL.Tag)
               Case "list"
                If COMM <> "" And LCase$(GetPiece(COMM, " ", 1)) <> "+ok" Then iTMP = iTMP + 1
        End Select
   End If
   If Mid(vtData, curPOS + 1, 1) = Chr$(10) Then curPOS = curPOS + 1
   If Mid(vtData, Len(vtData), 1) = Chr$(10) And curPOS >= bytesTotal - 1 Then Exit Sub
Loop Until curPOS >= bytesTotal
End Sub

Sub MainEmailOptions()
On Error Resume Next
emailCheckMail = False
Load frmSettings
frmSettings.CycleFrame 2
frmSettings.Show
frmSettings.txtUSER.SetFocus
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
          Dim Msg As Long
          Dim sFilter As String
          If ScaleMode <> 3 Then Msg = X / Screen.TwipsPerPixelX Else: Msg = X
          Select Case Msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK
                If wsock.State = sckClosed Or wsock.State = sckError And frmLogin.Visible = False Then
                    InitConnection
                Else
                    WindowState = 0
                    Visible = True
                    Show
                End If
             Case WM_RBUTTONDOWN
                mnuShowMessenger.Enabled = True
             Case WM_RBUTTONUP
                SetForegroundWindow hWnd
                DoEvents
                If mnuLogIn.Enabled Then
                    PopupMenu mnuPopup, , , , mnuLogIn
                Else
                    PopupMenu mnuPopup, , , , mnuShowMessenger
                End If
                
             Case WM_RBUTTONDBLCLK
          End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WindowState = vbMinimized Then
    Hide
Else
'If Me.Width < 3015 Then Me.Width = 3015
If Me.Height < 5000 Then Me.Height = 5000
If Me.Width < 3690 Then Me.Width = 3690
    'If Me.ScaleWidth - 240 >= 4095 Then
     '   tbToolBar.Width = Me.ScaleWidth - 240
    'End If
    lstUsers.ColumnHeaders(1).Width = 1800
    lnUp(1).X2 = ScaleWidth
    lnUp(0).X2 = ScaleWidth
    
    frameUsers.Top = frameToolBar.Top + frameToolBar.Height + 20
    frameUsers.Width = ScaleWidth
    frameUsers.Height = ScaleHeight - frameUsers.Top - cmdConnect.Height - 240 - picDown.Height - sbBar.Height
    
    
    lstUsers.Top = 120 + lblMyStatus.Top + lblMyStatus.Height
    lstUsers.Width = frameUsers.Width - 240
    lstUsers.Height = frameUsers.Height - lstUsers.Top - 840
    lstUsers.ColumnHeaders(2).Width = lstUsers.Width - lstUsers.ColumnHeaders(1).Width - 250
    
    lblWait.Top = lstUsers.Top + lstUsers.Height + 120
    lblWait.Left = 405
    lstWait.Top = lblWait.Top + lblWait.Height + 20
    lstWait.Height = frameUsers.Height - lstUsers.Top - lblWait.Height - lstUsers.Height - 240
    lstWait.Width = lstUsers.Width
    
    cmdConnect.Top = Me.ScaleHeight - cmdConnect.Height - 120 - picDown.Height - sbBar.Height
    
    lblName.Top = cmdConnect.Top - 40
    lblName.Left = Me.ScaleWidth - lblName.Width - 220
    
    imgStat.Top = lblName.Top - 150
    imgStat.Left = Me.ScaleWidth - imgStat.Width
    
    lblEmail.Top = imgStat.Top + imgStat.Height - 80
    lblEmail.Left = Me.ScaleWidth - lblEmail.Width - 120
    lstUsers.Refresh
    picConnecting.Left = 20
    picConnecting.Width = frameUsers.Width - 40
    picConnecting.Height = frameUsers.Height - lblMyStatus.Height - lblMyStatus.Top - 40
    picBannerHolder.Left = (ScaleWidth / 2) - (picBannerHolder.Width / 2)
    picBannerHolder.Top = (picDown.ScaleHeight / 2) - (picBannerHolder.Height / 2)
    'rX = (ScaleWidth / 2) - (frameToolBar.Width / 2)
    'If rX < 0 Then rX = 0
    'frameToolBar.Left = rX
    frameToolBar.Width = ScaleWidth
    DoEvents
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Not EndingApp Then
    Hide
    Cancel = True
Else
    DoEvents
    End
End If
End Sub


Private Sub lblEMail_Click()
On Error Resume Next
WebURL "http://www.vasilakis.com/"
End Sub

Private Sub lblMyStatus_Change()
On Error Resume Next
cmdStatus.Left = lblMyStatus.Left + lblMyStatus.Width + 30
End Sub

Private Sub lblMyStatus_Click()
On Error Resume Next
If cmdStatus.Enabled Then PopupMenu mnuStatus
mnuStatus.Visible = True
End Sub



Private Sub lstUsers_DblClick()
On Error Resume Next
Dim lstItem As ListItem
Dim rEmail As String
Dim rYN, rTemp
Set lstItem = lstUsers.SelectedItem
Err = 0
rEmail = lstItem.Text
If Err <> 0 Then Exit Sub
If rEmail = myEmail Then Exit Sub
If lstItem.Text = "Online" Then Exit Sub
If lstItem.Text = "Not Connected." Then Exit Sub
If lstItem.Text = "Not Online" Then Exit Sub
If lstItem.Text = "None." Then Exit Sub
If lstItem.SmallIcon = 1 Then
    If tmrAnimOnline.Enabled = True And lstItem.Index = 1 Then GoTo DoItAnyway
    rYN = MsgBox("User is currently offline. Would you like to send him/her an email?", vbYesNo + vbQuestion, "Offline User")
    If rYN = vbNo Then Exit Sub
    rTemp = ShellExecute(hWnd, "open", "mailto:" & rEmail, vbNullString, CurDir$, SW_SHOW)
    Exit Sub
End If
DoItAnyway:
If Not ExistsChat(lstItem.Text) Then
    CreateNewChat lstItem.Text, False
End If
End Sub


Private Sub lstUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim lstItem As ListItem
Dim rEmail As String

If Button = vbRightButton Then
    Set lstItem = lstUsers.SelectedItem
    Err = 0
    rEmail = lstItem.Text
    If lstItem = "Online" Then Exit Sub
    If lstItem = "Not Connected." Then Exit Sub
    If lstItem = "None." Then Exit Sub
    If lstItem = "Not Online" Then Exit Sub
    If Err <> 0 Then Exit Sub
        mnuSendFile.Enabled = True
        mnuDCC.Enabled = True
    If lstItem.SmallIcon = 1 Then
        If tmrAnimOnline.Enabled = True And lstItem.Index = 1 Then GoTo DoItAnyway
        mnuSendFile.Enabled = False
        mnuDCC.Enabled = False
    End If
DoItAnyway:
    PopupMenu mnuPopUser
End If
End Sub


Private Sub mnuClose_Click()
On Error Resume Next
Hide
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 1
'MsgBox "vasilakis Messenger v" & App.Major & "." & App.Minor & "." & App.Revision & ", created by vasilakis S.A." & _
 'vbCrLf & vbCrLf & "Contact ------------------------" & vbCrLf & _
 '"vasilakis S.A." & vbCrLf & vbCrLf & _
 '"ATHENS, Pl. Karysth 5, 10561" & _ vbcrlf
 '"LARISSA, Koyma 1, 41222" & vbCrLf & vbCrLf & "Tel: (+30) 41 537343" & vbCrLf & "Fax: (+30) 41 537343", vbInformation, "About vasilakis Messenger"
End Sub

Private Sub mnuBugReport_Click()
On Error Resume Next
frmBugReport.Show
End Sub


Private Sub mnuDCC_Click()
On Error Resume Next
Dim lstItem As ListItem
Dim rEmail As String

    Set lstItem = lstUsers.SelectedItem
    Err = 0
    rEmail = lstItem.Text
    If Err <> 0 Then Exit Sub
    CreateNewDCChat rEmail, True
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
If wsock.State = sckConnected Then
    rYN = MsgBox("You are currently Connected. Are you sure you want to quit?", vbYesNo + vbQuestion, "Exit vasilakis Messenger!")
    If rYN = vbYes Then
        EndApp
    End If
Else
    EndApp
End If
End Sub

Private Sub mnuLogIn_Click()
On Error Resume Next
InitConnection
End Sub

Private Sub mnuLogOut_Click()
On Error Resume Next
CloseConnection
End Sub


Private Sub mnuMySettings_Click()
On Error Resume Next
cmdToolBar_Click 3
End Sub

Private Sub mnuPopup_Click()
On Error Resume Next
If Me.Visible = True Then
    mnuShowMessenger.Enabled = False
End If
End Sub

Private Sub mnuSendFile_Click()
On Error Resume Next
Dim lstItem As ListItem
Dim rEmail As String

    Set lstItem = lstUsers.SelectedItem
    Err = 0
    rEmail = lstItem.Text
    If Err <> 0 Then Exit Sub
    CreateNewFileSend rEmail

End Sub


Private Sub mnuShowMessenger_Click()
On Error Resume Next
WindowState = 0
Show
End Sub

Private Sub mnuStatusAway_Click()
On Error Resume Next
SendData "status;1"
imgMyStatus.Picture = imgList.ListImages(4).Picture
lblMyStatus.Caption = myName & " (Away)"
TogleStatusMenu 1

End Sub

Private Sub mnuStatusBRB_Click()
On Error Resume Next
SendData "status;3"
imgMyStatus.Picture = imgList.ListImages(4).Picture
lblMyStatus.Caption = myName & " (Be Right Back)"
TogleStatusMenu 3
End Sub

Private Sub mnuStatusBusy_Click()
On Error Resume Next
SendData "status;2"
imgMyStatus.Picture = imgList.ListImages(4).Picture
lblMyStatus.Caption = myName & " (Busy)"
TogleStatusMenu 2
End Sub


Private Sub mnuStatusOnline_Click()
On Error Resume Next
SendData "status;0"
imgMyStatus.Picture = imgList.ListImages(2).Picture
lblMyStatus.Caption = myName & " (Online)"
TogleStatusMenu 0
End Sub


Private Sub mnuUserInfo_Click()
On Error Resume Next
Dim lstItem As ListItem
Dim rEmail As String

    Set lstItem = lstUsers.SelectedItem
    Err = 0
    rEmail = lstItem.Text
    If Err <> 0 Then Exit Sub
    SendData "userinfo;" & rEmail

End Sub

Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    picBannerHolder.Left = (ScaleWidth / 2) - (picBannerHolder.Width / 2)
    picBannerHolder.Top = (picDown.ScaleHeight / 2) - (picBannerHolder.Height / 2)
End Sub


Private Sub tmrActiveConnection_Timer()
On Error Resume Next
If ActiveConnection Then
    If cmdConnect.Picture = imgConnect.Picture Then
        tmrActiveConnection.Enabled = False
        InitConnection
    End If
End If
End Sub

Private Sub tmrAnimOnline_Timer()
On Error Resume Next
    ofSec = ofSec + 1
    mCounter = mCounter + 1
    lstUsers.ListItems.Item(2).SmallIcon = mCounter
    lstUsers.ListItems.Item(2).ForeColor = RGB(mCounter * 80, 0, 0)
    If mCounter > 2 Then: mCounter = 0
    If ofSec = 20 Then
        tmrAnimOnline.Enabled = False
        mCounter = 0
        lstUsers.ListItems.Item(2).SmallIcon = 2
        lstUsers.ListItems.Item(2).ForeColor = &H80000008
    End If
End Sub

Private Sub tmrBanner_Timer()
On Error Resume Next
CycleBanner
End Sub

Private Sub wsock_Close()
On Error Resume Next
rDialUp = GetSetting("vasilakis Messenger", "Login", "DialUp", "0")
If rDialUp = "1" And ActiveConnection = False Then
    tmrActiveConnection.Enabled = True
End If

wsock.Close
CloseConnection
End Sub

Private Sub wsock_Connect()
On Error Resume Next
Dim lstItem As ListItem
lstUsers.ListItems.Clear
Set lstItem = lstWait.ListItems.Add(1, , "None.")
Set lstItem = lstUsers.ListItems.Add(1, , "Online")
lstItem.bold = True
Set lstItem = lstUsers.ListItems.Add(2, , "None.")
Set lstItem = lstUsers.ListItems.Add(3, , "Not Online")
lstItem.bold = True
Set lstItem = lstUsers.ListItems.Add(4, , "None.")
lblName.Caption = "Authenticating..."
imgStat.Picture = imgOnline.Picture
cmdConnect.Picture = imgDisconnect.Picture
cmdConnect.ToolTipText = "Disconnect from host"
modIcon Me, IconObject.Handle, imgTrayOnLine.Picture, "vasilakis Messenger! - Online"
sbBar.SimpleText = "Connected! Authorizing..."
SendData "version;" & App.Major & "." & App.Minor & "." & App.Revision
mnuStatusOnline.Checked = True
'lstUsers.Enabled = True
End Sub


Private Sub wsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim lstItem As ListItem
Dim vtData As String
Dim curPOS As Single
Dim MESSAGE As String
Dim COMM As String
Dim rUser As String
Dim rMsg As String
Static rApp
Static iCount
Static rREC
wsock.GetData vtData
vtData = DecryptPassword(0.06, vtData)
wsock.Tag = wsock.Tag & vtData
GoTo CheckMSG
cmd:
    Select Case COMM
        Case "auth"
            lblMyStatus.Caption = "Logging In..."
            SendData "login;" & myEmail & vbLf & frmLogin.txtPASS.Text
        Case "noauth"
            SaveSetting "vasilakis Messenger", "Login", "SavePass", "0"
            DeleteSetting "vasilakis Messenger", "Login", "Password"
            MsgBox "Password not accepted.", vbCritical, "Error"
            CloseConnection
            Exit Sub
        Case "ping"
            SendData "pong"
        Case "yourname"
            myName = MESSAGE
            If JustLoggedIn = False Then
                lblName.Caption = myName
                If lblName.Caption = "" Then
                    lblName.Caption = "Online"
                Else
                    lblName.Caption = "Online"
                End If
            End If
            lblMyStatus.Caption = myName
        Case "login"
            sbBar.SimpleText = "Authorization Succesfull. Getting lists..."
            mnuLogOut.Enabled = True
            mnuStatus.Enabled = True
            imgMyStatus.Picture = imgList.ListImages(2).Picture
            cmdStatus.Enabled = True
            SaveSetting "vasilakis Messenger", "Login", "Email", myEmail
            SaveSetting "vasilakis Messenger", "Login", "Password", frmLogin.txtPASS.Text
            SaveSetting "vasilakis Messenger", "Login", "SavePass", frmLogin.chkSavePassword.Value
            Unload frmLogin
        Case "yourinfo"
            myYear = GetPiece(MESSAGE, vbLf, 1)
            mySex = GetPiece(MESSAGE, vbLf, 2)
            myCity = GetPiece(MESSAGE, vbLf, 3)
            myCountry = GetPiece(MESSAGE, vbLf, 4)
            
        Case "yourstatus"
            If Val(MESSAGE) = 0 Then
                imgMyStatus.Picture = imgList.ListImages(2).Picture
                lblMyStatus.Caption = myName & " (Online)"
            ElseIf Val(MESSAGE) = 1 Then
                imgMyStatus.Picture = imgList.ListImages(4).Picture
                lblMyStatus.Caption = myName & " (Away)"
            ElseIf Val(MESSAGE) = 2 Then
                imgMyStatus.Picture = imgList.ListImages(4).Picture
                lblMyStatus.Caption = myName & " (Busy)"
            ElseIf Val(MESSAGE) = 3 Then
                imgMyStatus.Picture = imgList.ListImages(4).Picture
                lblMyStatus.Caption = myName & " (Be Right Back)"
            End If
        Case "online"
            AddToList GetPiece(MESSAGE, vbLf, 1), True
            ChangeDisplayName GetPiece(MESSAGE, vbLf, 1), GetPiece(MESSAGE, vbLf, 2)
            If Not JustLoggedIn Then
                tmrAnimOnline.Enabled = True
                CreateNewPopupWindow GetPiece(MESSAGE, vbLf, 1)
            End If
        Case "offline"
            AddToList MESSAGE, False
            If ExistsChat(MESSAGE) Then
                AddServerChatText MESSAGE, "User went offline."
            End If
        Case "waitline"
            AddToWait MESSAGE
        Case "userstatus"
            ChangeUserStatus GetPiece(MESSAGE, vbLf, 1), Val(GetPiece(MESSAGE, vbLf, 2))
        Case "starthasyou"
            frmWhoHas.lstUsers.ListItems.Clear
        Case "userhasyou"
            Set lstItem = frmWhoHas.lstUsers.ListItems.Add(, , LCase$(GetPiece(MESSAGE, vbLf, 1)), , Val(GetPiece(MESSAGE, vbLf, 2)))
            lstItem.SubItems(1) = GetPiece(MESSAGE, vbLf, 3)
        Case "endhasyou"
            frmWhoHas.Show
            frmWhoHas.lstUsers.Refresh
        Case "declined"
            RemoveFromList MESSAGE
            MsgBox "User '" & MESSAGE & "' did not allow you to add him/her to your list.", vbCritical, "Declined"
        Case "addedblock"
            MsgBox "You've blocked user '" & MESSAGE & "' succesfully.", vbInformation, "Block List"
        Case "addeduser"
            AddToWait MESSAGE
            frmUserAdded.Show
        Case "nosuchuser"
            RemoveFromList MESSAGE
            MsgBox "User with email '" & MESSAGE & "', does not exist.", vbCritical, "Wrong Email Address"
        Case "displayname"
            ChangeDisplayName GetPiece(MESSAGE, vbLf, 1), GetPiece(MESSAGE, vbLf, 2)
        Case "userblockedyou"
            MsgBox "User '" & MESSAGE & "' have just added you to his/her block list." & vbCrLf & "You are unable to send to this user a message or see his/her online status unless he/she removes you from the block list.", vbInformation, "You are blocked."
            RemoveFromList LCase(MESSAGE)
        Case "userinfo"
            Unload frmUserinfo
            Load frmUserinfo
            frmUserinfo.txtName = GetNameFromEmail(GetPiece(MESSAGE, vbLf, 1))
            frmUserinfo.txtYear = GetPiece(MESSAGE, vbLf, 2)
            frmUserinfo.txtSex = GetPiece(MESSAGE, vbLf, 3)
            frmUserinfo.txtCity = GetPiece(MESSAGE, vbLf, 4)
            frmUserinfo.txtCountry = GetPiece(MESSAGE, vbLf, 5)
            frmUserinfo.Show 1
        Case "removeduser"
            RemoveFromList MESSAGE
        Case "useraddedyou"
            NewAuthorization MESSAGE
        Case "finishedlist"
            picConnecting.Visible = False
            lblName.Caption = myName
            If lblName.Caption = "" Then
                lblName.Caption = "Online"
            Else
                lblName.Caption = "Online"
            End If
            If CountOffline = 0 And lstUsers.ListItems(GetOffLineIndex + 1).Text <> "None." Then
                Set lstItem = lstUsers.ListItems.Add(GetOffLineIndex + 1, , "None.")
            End If
            sbBar.SimpleText = "Online."
            JustLoggedIn = False
            CycleBanner
        Case "incomingmsg"
            rUser = GetPiece(MESSAGE, vbLf, 1)
            rMsg = GetPiece(MESSAGE, vbLf, 2)
            If Not ExistsChat(rUser) Then
                NewChat rUser
                CreateNewChatPopupWindow rUser
            End If
            AddChatText rUser, rMsg
        Case "serverincomingmsg"
            rUser = GetPiece(MESSAGE, vbLf, 1)
            rMsg = GetPiece(MESSAGE, vbLf, 2)
            If ExistsChat(rUser) Then
                AddServerChatText rUser, rMsg
            End If
        Case "typing"
            ChangeChatStatus MESSAGE, True
        Case "notyping"
            ChangeChatStatus MESSAGE, False
        Case "closedchat"
            ChangeChatStatusBar MESSAGE, GetNameFromEmail(MESSAGE) & " has closed the chat window."
        Case "servermessage"
            MsgBox MESSAGE, vbInformation, "Server Message"
        Case "dcc"
            CreateNewDCChat GetPiece(MESSAGE, vbLf, 1), False, GetPiece(MESSAGE, vbLf, 2), GetPiece(MESSAGE, vbLf, 3)
        Case "sendfiles"
            CreateNewFileReceive GetPiece(MESSAGE, vbLf, 1), GetPiece(MESSAGE, vbLf, 2), GetPiece(MESSAGE, vbLf, 3)
   
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

Sub SendData(Text As String)
On Error Resume Next
If wsock.State <> sckConnected Then CloseConnection: Exit Sub

wsock.SendData EncryptPassword(0.06, Text & vbCrLf)
DoEvents
End Sub

Private Sub wsockBanner_Connect()
On Error Resume Next
wsockBanner.Tag = ""
End Sub

Private Sub wsockBanner_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim vtData As String
Dim curPOS As Single
Dim COMM As String
Dim MESSAGE As String
Dim ws As Winsock
Randomize Timer
curPOS = 0
wsockBanner.GetData vtData
If vtData = "" Then Exit Sub
Select Case wsockBanner.Tag
    Case ""
        URL = vtData
        wsockBanner.Tag = "SIZE"
        wsockBanner.SendData "SIZE"
    Case "SIZE"
        bannerSize = Val(vtData)
        bannercurSize = 0
        bannerdata = ""
        wsockBanner.Tag = "FILE"
        wsockBanner.SendData "FILE"

    Case "FILE"
        bannerdata = bannerdata & vtData
        bannercurSize = bannercurSize + Len(vtData)
        If bannercurSize >= bannerSize Then
            LoadAniGif AnimatedGIF
            wsockBanner.Close
        End If
        CenterBanners
End Select
DoEvents
End Sub
