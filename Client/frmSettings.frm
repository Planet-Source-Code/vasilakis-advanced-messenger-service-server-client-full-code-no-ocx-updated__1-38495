VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MySettings"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
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
   ScaleHeight     =   3855
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdFrame 
      BackColor       =   &H00D3D3C3&
      Caption         =   "Various Settings"
      Height          =   495
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdFrame 
      BackColor       =   &H00D3D3C3&
      Caption         =   "Connection Settings"
      Height          =   495
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFrame 
      BackColor       =   &H00D3D3C3&
      Caption         =   "My Info"
      Height          =   495
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame frameSettings 
      BackColor       =   &H00D3D3C3&
      Height          =   2535
      Index           =   2
      Left            =   720
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtPOP 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtPASS 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtUSER 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkAutoCheckMail 
         BackColor       =   &H00D3D3C3&
         Caption         =   "Auto Check my email."
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkStartUp 
         BackColor       =   &H00D3D3C3&
         Caption         =   "Start Messenger when my Computer starts."
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POP3 Server"
         Enabled         =   0   'False
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Enabled         =   0   'False
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Enabled         =   0   'False
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Preferences..."
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
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Width           =   1590
      End
   End
   Begin VB.Frame frameSettings 
      BackColor       =   &H00D3D3C3&
      Height          =   2535
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   4455
      Begin VB.OptionButton optDialUp 
         BackColor       =   &H00D3D3C3&
         Caption         =   "... is a Dial Up Connection."
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optLAN 
         BackColor       =   &H00D3D3C3&
         Caption         =   "... is a Local Area Network Connection."
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3C3&
         Caption         =   "Please choose exactly the type of the connection you are using, because it will make the Messenger more reliable."
         Height          =   975
         Left            =   720
         TabIndex        =   8
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Line lnUp 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   720
         X2              =   3840
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Connection..."
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
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   1515
      End
      Begin VB.Line lnUp 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   720
         X2              =   3840
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame frameSettings 
      BackColor       =   &H00D3D3C3&
      Height          =   2535
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.Line lnUp 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   360
         X2              =   3480
         Y1              =   840
         Y2              =   840
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
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   915
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
         Left            =   360
         TabIndex        =   4
         Top             =   960
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
         Left            =   2040
         TabIndex        =   3
         Top             =   960
         Width           =   315
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
         Left            =   360
         TabIndex        =   2
         Top             =   1680
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
         Left            =   2040
         TabIndex        =   1
         Top             =   1680
         Width           =   330
      End
      Begin VB.Line lnUp 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   360
         X2              =   3480
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSettings.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CycleFrame(Index As Integer)
For i = 0 To cmdFrame.Count - 1
    If i <> Index Then
        frameSettings(i).Visible = False
        cmdFrame(i).FontBold = False
    End If
Next i
    frameSettings(Index).Visible = True
    cmdFrame(Index).FontBold = True
    If Index = 0 Then
        txtName.SetFocus
    ElseIf Index = 1 Then
        If optDialUp.Value = True Then optDialUp.SetFocus Else Me.optLAN.SetFocus
    ElseIf Index = 2 Then
        chkStartUp.SetFocus
    End If
End Sub

Private Sub chkAutoCheckMail_Click()
If chkAutoCheckMail.Value = 0 Then
   lblLabel(0).Enabled = False
   lblLabel(1).Enabled = False
   lblLabel(2).Enabled = False
   txtUSER.Enabled = False
   txtPASS.Enabled = False
   txtPOP.Enabled = False
Else
   lblLabel(0).Enabled = True
   lblLabel(1).Enabled = True
   lblLabel(2).Enabled = True
   txtUSER.Enabled = True
   txtPASS.Enabled = True
   txtPOP.Enabled = True
End If
End Sub

Private Sub cmdAccept_Click()
On Error Resume Next
DoEvents
Hide
rAutoCheckMail = GetSetting("OmniCom Messenger", "Settings", "AutoCheckMail", "0")
If chkAutoCheckMail.Value = 1 Then
    SaveSetting "OmniCom Messenger", "Settings", "AutoCheckMail", "1"
    SaveSetting "OmniCom Messenger", "Settings", "emailUsername", txtUSER.Text
    SaveSetting "OmniCom Messenger", "Settings", "emailPassword", EncryptPassword(0.06, txtPASS.Text)
    SaveSetting "OmniCom Messenger", "Settings", "emailPOPServer", txtPOP.Text
    frmMain.emailCheckMail = True
    frmMain.emailUsername = txtUSER.Text
    frmMain.emailPassword = txtPASS.Text
    frmMain.emailPOPServer = txtPOP.Text
    frmMain.CheckMail
Else
    frmMain.emailCheckMail = False
    frmMain.tmrCheckEmail.Enabled = False
    SaveSetting "OmniCom Messenger", "Settings", "AutoCheckMail", "0"
    DeleteSetting "OmniCom Messenger", "Settings", "emailUsername"
    DeleteSetting "OmniCom Messenger", "Settings", "emailPassword"
End If

If chkStartUp.Value = 1 Then
    SaveSetting "OmniCom Messenger", "Settings", "OnStartUp", "1"
    RegistryCreateNewKey rrkHKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run"
    RegistrySetKeyValue rrkHKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run", "OmniCom Messenger", App.Path & "\" & App.EXEName & ".exe", rrkRegSZ
Else
    SaveSetting "OmniCom Messenger", "Settings", "OnStartUp", "0"
    RegistryCreateNewKey rrkHKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run"
    RegistrySetKeyValue rrkHKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run", "OmniCom Messenger", "", rrkRegSZ
End If
If optDialUp.Value = True Then
    SaveSetting "OmniCom Messenger", "Login", "DialUp", "1"
    DeleteSetting "OmniCom Messenger", "Login", "LAN"
    If frmMain.wsock.State <> sckConnected Then
        frmMain.tmrActiveConnection.Enabled = True
    End If
Else
    SaveSetting "OmniCom Messenger", "Login", "LAN", "1"
    DeleteSetting "OmniCom Messenger", "Login", "DialUp"
    frmMain.tmrActiveConnection.Enabled = False
End If
If frmMain.wsock.State = sckConnected Then
    If txtName.Text <> frmMain.myName Then
        frmMain.SendData "changemyname;" & txtName.Text
    End If
    frmMain.SendData "myinfo;" & cboYear.Text & vbLf & cboSex.Text & vbLf & txtCity.Text & vbLf & txtCountry.Text
End If
Unload Me
End Sub


Private Sub cmdCancel_Click()
Unload Me
DoEvents
End Sub


Private Sub cmdFrame_Click(Index As Integer)
CycleFrame Index
End Sub

Private Sub Form_Activate()
SetWindowPos hwnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i
rAutoCheckMail = GetSetting("OmniCom Messenger", "Settings", "AutoCheckMail", "0")
If rAutoCheckMail = "1" Then
    chkAutoCheckMail.Value = 1
    txtUSER.Text = GetSetting("OmniCom Messenger", "Settings", "emailUsername")
    txtPASS.Text = DecryptPassword(0.06, GetSetting("OmniCom Messenger", "Settings", "emailPassword"))
    txtPOP.Text = GetSetting("OmniCom Messenger", "Settings", "emailPOPServer")
Else
    chkAutoCheckMail.Value = 0
End If
chkAutoCheckMail_Click
rOnStartUp = GetSetting("OmniCom Messenger", "Settings", "OnStartUp", "0")
If rOnStartUp = "1" Then
    chkStartUp.Value = 1
Else
    chkStartUp.Value = 0
End If

rDialUp = GetSetting("OmniCom Messenger", "Login", "DialUp", "0")
If rDialUp = "1" Then
    optDialUp.Value = True
Else
    optLAN.Value = True
End If
For i = 1933 To Val(Year(Now))
    cboYear.AddItem Trim(Str$(i))
Next i
cboSex.AddItem "M"
cboSex.AddItem "F"
End Sub


