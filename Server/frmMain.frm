VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vasilakis :messenger Server"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
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
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
   End
   Begin VB.ListBox lstUsers 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Timer tmrCheckList 
      Interval        =   1000
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Timer tmrUser 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   3000
      Top             =   1920
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3135
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   2520
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "BroadCast Message"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Label lblBytes 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 bytes"
      Height          =   195
      Left            =   3555
      TabIndex        =   5
      Top             =   960
      Width           =   540
   End
   Begin VB.Label llbInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Traffic"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Users online."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IconObject As Object

Public bytesReceived As Double
Function AddAllowUserList(BelongsTo As String, who As String) As Boolean
On Error Resume Next
Dim rs As New ADODB.Recordset
'rs.open "allowlists")
rs.Open "allowlists", db, adOpenKeyset, adLockOptimistic
With rs
    .AddNew
    !UserName = BelongsTo
    !Addedusername = who
    !Auth = False
    Err = 0
    .Update
    If Err <> 0 Then
        AddAllowUserList = False
        Exit Function
    End If
End With
AddAllowUserList = True
End Function

Function AddBlockUserList(BelongsTo As String, who As String) As Boolean
On Error Resume Next
Dim rs As New ADODB.Recordset
'rs.open "allowlists")
rs.Open "blocklists", db, adOpenKeyset, adLockOptimistic
With rs
    .AddNew
    !UserName = BelongsTo
    !Blockedusername = who
    .Update
End With
AddBlockUserList = True
End Function


Sub AddUserToLOGList(eMail As String)
Dim i
For i = 0 To lstUsers.ListCount - 1
    If LCase(lstUsers.List(i)) = LCase(eMail) Then Exit Sub
Next i
lstUsers.AddItem eMail, 0
DoEvents
End Sub

Sub RemoveUserFromLOGList(eMail As String)
Dim i
For i = 0 To lstUsers.ListCount - 1
    If LCase(lstUsers.List(i)) = LCase(eMail) Then
        lstUsers.RemoveItem i
    End If
Next i
DoEvents
End Sub

Sub ChangeAuthorization(Index As Integer, who As String, Allow As Byte)
On Error Resume Next
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM allowlists WHERE Username = '" & who & "' AND Addedusername = '" & Clients(Index).User & "';", db, adOpenKeyset, adLockOptimistic
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
If Allow = 1 Then
    'rs.Edit
    rs!Auth = True
    rs.Update
    If IsOnline(who) Then
        SendData GetIndexFromUsername(who), "online;" & Clients(Index).User & vbLf & GetUserDisplayName(Clients(Index).User)
    End If
Else
    rs.Delete adAffectCurrent
    rs.UpdateBatch
    If IsOnline(who) Then
        SendData GetIndexFromUsername(who), "declined;" & Clients(Index).User
    End If
End If

End Sub

Sub ChangeUserInfo(Index As Integer, rYear As String, rSex As String, rCity As String, rCountry As String)

Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM users WHERE Username = '" & Clients(Index).User & "';", db, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If Err <> 0 Then Exit Sub
'rs.edit
rs!Year = rYear
rs!Sex = rSex
rs!City = rCity
rs!Country = rCountry
rs.Update
SendUserInfo Index
End Sub

Sub ChangeUserStatus(Index As Integer, rAWAY As AWAYSTATUS)
On Error Resume Next
Dim rUser As String
Clients(Index).Status = rAWAY
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
Do While Not rs.EOF
    If IsOnline(rs.Fields("Username")) And rs.Fields("Auth") Then
        rUser = GetUserDisplayName(Clients(Index).User)
        If rUser <> "" Then
            SendData GetIndexFromUsername(rs.Fields("username")), "displayname;" & Clients(Index).User & vbLf & rUser
        End If

        SendData GetIndexFromUsername(rs.Fields("username")), "userstatus;" & Clients(Index).User & vbLf & Clients(Index).Status
    End If
    rs.MoveNext
Loop
End Sub

Function ClientsCount() As Integer
ClientsCount = 0
Dim ws As Winsock
For Each ws In wsock
    If ws.State = sckConnected And ws.Index > 0 Then
        ClientsCount = ClientsCount + 1
    End If
Next
End Function

Function RemoveAllowUserList(BelongsTo As String, who As String) As Boolean
On Error Resume Next
Dim rs As New ADODB.Recordset

rs.Open "SELECT * FROM allowlists WHERE Username = '" & BelongsTo & "' AND Addedusername = '" & who & "';", db, adOpenKeyset, adLockOptimistic
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Function
rs.Delete adAffectCurrent
rs.UpdateBatch
RemoveAllowUserList = True
End Function


Function RemoveBlockUserList(BelongsTo As String, who As String) As Boolean
On Error Resume Next
Dim rs As New ADODB.Recordset

rs.Open "SELECT * FROM blocklists WHERE Username = '" & BelongsTo & "' AND Blockedusername = '" & who & "';", db, adOpenKeyset, adLockOptimistic
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Function
rs.Delete adAffectCurrent
rs.UpdateBatch
RemoveBlockUserList = True
End Function



Sub CloseClient(Index As Integer)
On Error Resume Next
RemoveUserFromLOGList Clients(Index).User
SendUserIsOffline Index
Clients(Index).Auth = False
Clients(Index).User = ""
Clients(Index).iCount = 0
Clients(Index).rREC = 0
Clients(Index).Pong = False
tmrUser(Index).Enabled = False
Unload tmrUser(Index)
wsock(Index).Close
wsock(Index).Tag = ""
Unload wsock(Index)
sbBar.SimpleText = ClientsCount & " Clients connected."
End Sub

Function GetIndexFromUsername(rUser As String) As Integer
On Error Resume Next
Dim ws As Winsock
For Each ws In wsock
    If ws.Index > 0 Then
        If Clients(ws.Index).User = rUser Then
            GetIndexFromUsername = ws.Index
            Exit Function
        End If
    End If
Next
GetIndexFromUsername = 0
End Function

Function GetUserDisplayName(rUser As String) As String
On Error Resume Next
DoEvents
GetUserDisplayName = ""
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM users WHERE Username = '" & rUser & "';", db
    Err = 0
    GetUserDisplayName = rs.Fields("DisplayName")
    If Err <> 0 Then GetUserDisplayName = rUser

End Function

Function IsOnline(who As String) As Boolean
Dim rTMP
rTMP = GetIndexFromUsername(who)
If rTMP = 0 Then IsOnline = False Else IsOnline = True
End Function

Sub SendAllowList(Index As Integer)
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim UserInList As String
    rs.Open "SELECT * FROM allowlists WHERE Username = '" & Clients(Index).User & "' ORDER BY Addedusername DESC;", db
    Err = 0
    rs.MoveFirst
    If Err <> 0 Then SendData Index, "finishedlist": Exit Sub
    Do While Not rs.EOF
        UserInList = rs.Fields("Addedusername")
        If rs.Fields("Auth") Then
            If IsOnline(UserInList) Then
                SendData Index, "online;" & UserInList
                SendData Index, "userstatus;" & UserInList & vbLf & Clients(GetIndexFromUsername(UserInList)).Status
            Else
                SendData Index, "offline;" & UserInList
            End If
        Else
            SendData Index, "waitline;" & UserInList
        End If
        rs.MoveNext
    Loop
    SendData Index, "finishedlist"
End Sub

Sub SendUserInfo(Index As Integer)
On Error Resume Next
Dim rYear As String, rSex As String, rCity As String, rCountry As String
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM users WHERE Username = '" & Clients(Index).User & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
rYear = rs.Fields("year")
rSex = rs.Fields("sex")
rCity = rs.Fields("city")
rCountry = rs.Fields("country")
SendData Index, "yourinfo;" & rYear & vbLf & rSex & vbLf & rCity & vbLf & rCountry
End Sub

Sub SendToUserInfo(Index As Integer, who As String)
On Error Resume Next
Dim rYear As String, rSex As String, rCity As String, rCountry As String
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM users WHERE Username = '" & who & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
rYear = rs.Fields("year")
rSex = rs.Fields("sex")
rCity = rs.Fields("city")
rCountry = rs.Fields("country")
SendData Index, "userinfo;" & who & vbLf & rYear & vbLf & rSex & vbLf & rCity & vbLf & rCountry
End Sub


Sub SendWhoHasYou(Index As Integer)
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim ws As Winsock
rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
Err = 0
SendData Index, "starthasyou;"
rs.MoveFirst
If Err <> 0 Then GoTo HasBlocked
Do While Not rs.EOF
    SendData Index, "userhasyou;" & rs.Fields("Username") & vbLf & IIf(IsOnline(rs.Fields("Username")), "2", "1") & vbLf & GetUserDisplayName(rs.Fields("Username"))
    rs.MoveNext
Loop
HasBlocked:
rs.Close
rs.Open "SELECT * FROM blocklists WHERE Username = '" & Clients(Index).User & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then SendData Index, "endhasyou;": Exit Sub
Do While Not rs.EOF
    SendData Index, "userhasyou;" & rs.Fields("blockedUsername") & vbLf & "4" & vbLf & GetUserDisplayName(rs.Fields("blockedUsername"))
    rs.MoveNext
Loop
SendData Index, "endhasyou;"
End Sub


Sub CheckAllowList(Index As Integer)
On Error Resume Next

Dim rs As New ADODB.Recordset
Dim UserInList As String
    rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
    Err = 0
    rs.MoveFirst
    If Err <> 0 Then Exit Sub
    Do While Not rs.EOF
        UserInList = rs.Fields("Addedusername")
        If Not rs.Fields("Auth") Then
            SendData Index, "useraddedyou;" & rs.Fields("Username")
        End If
        rs.MoveNext
    Loop

End Sub

Sub SendData(Index As Integer, Text As String)
On Error Resume Next
Err = 0
wsock(Index).SendData EncryptPassword(0.06, Text & vbCrLf)
DoEvents
If Err <> 0 Then CloseClient Index
End Sub

Sub SendUserIsOnline(Index As Integer)
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim ws As Winsock
rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
Do While Not rs.EOF
    If IsOnline(rs.Fields("Username")) And rs.Fields("Auth") Then
        SendData GetIndexFromUsername(rs.Fields("Username")), "online;" & Clients(Index).User & vbLf & GetUserDisplayName(Clients(Index).User)
    End If
    rs.MoveNext
Loop
End Sub

Sub SendNewNameToLists(Index As Integer, NName As String)
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim ws As Winsock
rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
Err = 0
rs.MoveFirst
If Err <> 0 Then Exit Sub
Do While Not rs.EOF
    If IsOnline(rs.Fields("Username")) Then
        SendData GetIndexFromUsername(rs.Fields("Username")), "displayname;" & Clients(Index).User & vbLf & NName
        SendData GetIndexFromUsername(rs.Fields("username")), "userstatus;" & Clients(Index).User & vbLf & Clients(Index).Status
    End If
    rs.MoveNext
Loop
End Sub


Sub SendUserIsOffline(Index As Integer)

On Error Resume Next
Dim rs As New ADODB.Recordset
Dim ws As Winsock
rs.Open "SELECT * FROM allowlists WHERE Addedusername = '" & Clients(Index).User & "';", db
Err = 0
'rs.MoveFirst
If rs.EOF Then Exit Sub
Do While Not rs.EOF
    If IsOnline(rs.Fields("Username")) And rs.Fields("Auth") Then
        SendData GetIndexFromUsername(rs.Fields("Username")), "offline;" & Clients(Index).User
    End If
    rs.MoveNext
Loop
End Sub

Function UserExistsInDB(rUser As String) As Boolean
On Error Resume Next
UserExistsInDB = False
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM users WHERE Username = '" & rUser & "';", db
    Err = 0
    rs.MoveFirst
    If Err <> 0 Then UserExistsInDB = False: Exit Function
    UserExistsInDB = True
End Function

Function IsBlocked(who As String, whoBlocked As String) As Boolean
On Error Resume Next
IsBlocked = False
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM blocklists WHERE Username = '" & who & "' AND blockedusername = '" & whoBlocked & "';", db
    Err = 0
    rs.MoveFirst
    If Err <> 0 Then IsBlocked = False: Exit Function
    IsBlocked = True
End Function

Function ValidateUser(rUser As String, rPass As String) As Boolean
On Error Resume Next
ValidateUser = False
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM users WHERE Username = '" & rUser & "';", db
    Err = 0
'    rs.MoveFirst
    If rs.EOF Then ValidateUser = False: Exit Function
    If rPass = rs.Fields("Password") Then ValidateUser = True: Exit Function
    ValidateUser = False
End Function

Private Sub cmdHide_Click()
Hide
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
Dim ws As Winsock
Dim i
i = 0
If GetPiece(txtSend.Text, " ", 1) <> "" Then
    For Each ws In wsock
        If ws.State = sckConnected And ws.Index <> 0 Then
            i = i + 1
            SendData ws.Index, "servermessage;" & txtSend.Text
        End If
    Next ws
MsgBox "Message was sent to " & i & " client(s).", vbInformation, "BroadCast Message"
txtSend.Text = ""
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then End
Set IconObject = Me.Icon
AddIcon Me, IconObject.Handle, IconObject, "vasilakis Messenger [Server]"
wsock(0).RemotePort = 0
wsock(0).LocalPort = "9811"
wsock(0).Listen
sbBar.SimpleText = ClientsCount & " Clients connected."
Visible = False
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
                WindowState = 0
                Show
             Case WM_RBUTTONDOWN
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Sub


Private Sub Form_Resize()
If Me.WindowState = 1 Then Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
delIcon IconObject.Handle
End
'Cancel = True
End Sub




Private Sub tmrCheckList_Timer()
If wsock(0).State <> sckListening Then
    wsock(0).Close
    wsock(0).Listen
End If
End Sub

Private Sub tmrUser_Timer(Index As Integer)
If Index = 0 Then Exit Sub
If Clients(Index).Pong = False Then
    CloseClient Index
    Exit Sub
End If
Clients(Index).Pong = False
SendData Index, "ping;"
End Sub

Private Sub wsock_Close(Index As Integer)
On Error Resume Next
If Index <> 0 Then
    CloseClient Index
Else
    wsock(0).Listen
End If

End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim ic As Integer
Dim ws As Winsock
ic = 0
Do
ic = ic + 1
 Err = 0
 Load wsock(ic)
 DoEvents
 If wsock(ic).State <> sckConnected And wsock(ic).State <> sckConnecting Then
    DoEvents
    wsock(ic).Close
    wsock(ic).Accept requestID
    ReDim Preserve Clients(wsock.Count + 1)
    Load tmrUser(ic)
    tmrUser(ic).Interval = 60000
    tmrUser(ic).Enabled = False
    Clients(ic).Pong = True
    Clients(ic).Auth = False
    Clients(ic).User = ""
    Clients(ic).iCount = 0
    Clients(ic).rREC = ""
    Clients(ic).Status = msgNONE
    sbBar.SimpleText = ClientsCount & " Clients connected."
    SendData ic, "auth;"
    DoEvents
  Exit Sub
 End If
Loop

End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim vtData As String
Dim curPOS As Single
Dim rName As String
Dim rUser As String
Dim rPass As String
Dim rPort As Double
Dim rMSG As String
Dim rTemp As String
Dim rIndex As Integer
Dim iTemp, rTempV, rTMP As String
Dim rREC, Comm, Message As String
Dim rFiles As String
Dim myYear As String, mySex As String, myCity As String, myCountry As String
Dim ws As Winsock
Randomize Timer
curPOS = 0
wsock(Index).GetData vtData
bytesReceived = bytesReceived + bytesTotal
lblBytes.Caption = bytesReceived & "bytes"
vtData = DecryptPassword(0.06, vtData)
wsock(Index).Tag = wsock(Index).Tag & vtData

GoTo CheckMSG
cmd:
curPOS = 0
    

If Clients(Index).Auth = False Then
        Select Case Comm
            Case "version"
                Clients(Index).Version = Message
            Case "login"
                If Clients(Index).Version = "" Then
                    SendData Index, "servermessage;Your messenger BETA Version has expired. Please upgrade from the site http://www.vasilakis.com/. Please support this greek service by using it and telling your friends! Thanks a lot!"
                    CloseClient Index
                ElseIf Clients(Index).Version = "1.2.0" Then
                    SendData Index, "servermessage;Your messenger BETA Version has expired!. Please upgrade from the site http://www.vasilakis.com/. Please support this greek service by using it and telling your friends! Thanks a lot!"
                    CloseClient Index
                ElseIf Clients(Index).Version = "1.5.0" Then
                    SendData Index, "servermessage;Your messenger BETA Version has expired!. Please upgrade from the site http://www.vasilakis.com/. Please support this greek service by using it and telling your friends! Thanks a lot!"
                    CloseClient Index
                ElseIf Clients(Index).Version = "1.6.5" Then
                    SendData Index, "servermessage;There is a new version available to download!. Please upgrade from the site http://www.vasilakis.com/. Please support this greek service by using it and telling your friends! Thanks a lot!"
                End If
                rUser = LCase$(GetPiece(Message, vbLf, 1))
                rPass = GetPiece(Message, vbLf, 2)
                If ValidateUser(rUser, rPass) Then
                    If IsOnline(rUser) Then
                        SendData GetIndexFromUsername(rUser), "servermessage;Your connection has been closed, because you have logged in from onother location."
                        CloseClient GetIndexFromUsername(rUser)
                    End If
                    DoEvents
                    rName = GetUserDisplayName(rUser)
                    SendData Index, "login;"
                    SendData Index, "yourname;" & rName
                    SendData Index, "yourstatus;" & Clients(Index).Status
                    Clients(Index).User = rUser
                    Clients(Index).Auth = True
                    SendAllowList Index
                    SendUserIsOnline Index
                    CheckAllowList Index
                    SendUserInfo Index
                    tmrUser(Index).Enabled = True
                    RemoveUserFromLOGList rUser
                    AddUserToLOGList rUser
                Else
                    SendData Index, "noauth;"
                    CloseClient Index
                    Exit Sub
                End If
            Case Else
                SendData Index, "noauth"
                CloseClient Index
                Exit Sub
        End Select
    Else
        Select Case Comm
            Case "pong"
                Clients(Index).Pong = True
            Case "list"
                Select Case Message
                    Case "new"
                    Case "all"
                    Case "old"
                    Case "users"
                End Select
            Case "changemyname"
                If Message <> "" Then
                    NewName Index, Message
                    rName = GetUserDisplayName(Clients(Index).User)
                    SendData Index, "yourname;" & Message
                    SendData Index, "yourstatus;" & Clients(Index).Status
                End If
            Case "myinfo"
                myYear = GetPiece(Message, vbLf, 1)
                mySex = GetPiece(Message, vbLf, 2)
                myCity = GetPiece(Message, vbLf, 3)
                myCountry = GetPiece(Message, vbLf, 4)
                ChangeUserInfo Index, myYear, mySex, myCity, myCountry
            Case "userinfo"
                If Not IsBlocked(Message, Clients(Index).User) Then
                    SendToUserInfo Index, Message
                End If
            Case "unblockuser"
                Message = LCase$(Message)
                    If RemoveBlockUserList(Clients(Index).User, Message) Then
                        SendData Index, "unblocked;" & Message
                    End If
            Case "blockuser"
                Message = LCase$(Message)
                If UserExistsInDB(Message) Then
                    If AddBlockUserList(Clients(Index).User, Message) Then
                        SendData Index, "addedblock;" & Message
                        RemoveAllowUserList Message, Clients(Index).User
                        If IsOnline(Message) Then
                            SendData GetIndexFromUsername(Message), "userblockedyou;" & Clients(Index).User
                        End If
                    End If
                Else
                    SendData Index, "nosuchuser;" & Message
                End If
            Case "adduser"
                Message = LCase$(Message)
                If UserExistsInDB(Message) Then
                    If Not IsBlocked(Message, Clients(Index).User) Then
                        If AddAllowUserList(Clients(Index).User, Message) Then
                            SendData Index, "addeduser;" & Message
                            If IsOnline(Message) Then
                                SendData GetIndexFromUsername(Message), "useraddedyou;" & Clients(Index).User
                            End If
                        End If
                    Else
                        SendData Index, "servermessage;You cannot add this user to your list, because he/she has blocked you."
                    End If
                Else
                    SendData Index, "nosuchuser;" & Message
                End If
            Case "getname"
                rUser = GetUserDisplayName(Message)
                If rUser <> "" Then
                    SendData Index, "displayname;" & Message & vbLf & rUser
                End If
            Case "removefromallowlist"
                If Message <> "" Then
                    RemoveAllowUserList Clients(Index).User, Message
                    SendData Index, "removeduser;" & Message
                End If
            Case "allowlist"
                ChangeAuthorization Index, GetPiece(Message, vbLf, 1), Val(GetPiece(Message, vbLf, 2))
            Case "whohasme"
                SendWhoHasYou Index
            Case "outgoingmessage"
                    If Not IsBlocked(GetPiece(Message, vbLf, 1), Clients(Index).User) Then
                        rUser = GetPiece(Message, vbLf, 1)
                        rMSG = GetPiece(Message, vbLf, 2)
                        SendChatMessage Index, rUser, rMSG
                    Else
                        SendData Index, "serverincomingmsg;" & GetPiece(Message, vbLf, 1) & vbLf & "You cannot send a message to this user, because he/she has blocked you. This user did NOT receive your message."
                    End If

            Case "typing"
                rIndex = GetIndexFromUsername(Message)
                SendData rIndex, "typing;" & Clients(Index).User
            Case "notyping"
                rIndex = GetIndexFromUsername(Message)
                SendData rIndex, "notyping;" & Clients(Index).User
            Case "closedchat"
                rIndex = GetIndexFromUsername(Message)
                SendData rIndex, "closedchat;" & Clients(Index).User
            Case "dcc"
                If IsOnline(GetPiece(Message, vbLf, 1)) Then
                    If Not IsBlocked(Message, Clients(Index).User) Then
                        rIndex = GetIndexFromUsername(GetPiece(Message, vbLf, 1))
                        rPort = Val(GetPiece(Message, vbLf, 2))
                        SendData rIndex, "dcc;" & Clients(Index).User & vbLf & wsock(Index).RemoteHostIP & vbLf & rPort
                    Else
                        SendData Index, "servermessage;You cannot request chat from this user, because he/she has blocked you."
                    End If
                End If
            Case "sendfiles"
                If IsOnline(GetPiece(Message, vbLf, 1)) Then
                    If Not IsBlocked(Message, Clients(Index).User) Then
                        rIndex = GetIndexFromUsername(GetPiece(Message, vbLf, 1))
                        rPort = Val(GetPiece(Message, vbLf, 2))
                        rFiles = GetPiece(Message, vbLf, 3)
                        SendData rIndex, "sendfiles;" & Clients(Index).User & vbLf & wsock(Index).RemoteHostIP & vbLf & rPort & vbLf & rFiles
                    Else
                        SendData Index, "servermessage;You cannot request a file transfer mode from this user, because he/she has blocked you."
                    End If
                End If
            Case "status"
                If Val(Message) <> Clients(Index).Status Then
                    ChangeUserStatus Index, Val(Message)
                End If
            End Select
    End If
NoCom:
CheckMSG:
    rTemp = wsock(Index).Tag
    If rTemp = "" Then Exit Sub
    
    Do
        Clients(Index).iCount = Clients(Index).iCount + 1
        
        iTemp = Mid(rTemp, Clients(Index).iCount, 1)
        rTempV = wsock(Index).Tag
        'vtData = Right(rTEMP, Clients(index).iCount)
        If Mid(rTemp, Clients(Index).iCount, 2) = vbCrLf Then
                wsock(Index).Tag = Right(rTemp, Len(rTemp) - (Clients(Index).iCount + 1))
                Clients(Index).iCount = 0
                Comm = ""
                Message = ""
                curPOS = 0
                Do
                    curPOS = curPOS + 1
                    rTMP = Left(Clients(Index).rREC, curPOS)
                    rTMP = Right(rTMP, 1)
                    If rTMP = ";" Then Exit Do
                    Comm = Comm & rTMP
                    'DoEvents
                Loop Until curPOS >= Len(Clients(Index).rREC)
                Comm = LCase$(Comm)
                Do
                    If curPOS = Len(Clients(Index).rREC) Then Exit Do
                    curPOS = curPOS + 1
                    rTMP = Left(Clients(Index).rREC, curPOS)
                    rTMP = Right(rTMP, 1)
                    Message = Message & rTMP
                    'Doevents
                Loop Until curPOS >= Len(Clients(Index).rREC)
                rTemp = ""
                Clients(Index).rREC = ""
                GoTo cmd
            Else
                Clients(Index).rREC = Clients(Index).rREC & iTemp
                DoEvents
        End If
   Loop Until wsock(Index).Tag = "" Or Clients(Index).iCount >= Len(wsock(Index).Tag)
DoEvents
End Sub

Sub SendChatMessage(Index As Integer, who As String, Msg As String)
Dim rIndex As Integer
rIndex = GetIndexFromUsername(who)
If rIndex = 0 Then
    SendData Index, "serverincomingmsg;" & who & vbLf & "User did not receive your message because he/she is offline."
    Exit Sub
End If
SendData rIndex, "incomingmsg;" & Clients(Index).User & vbLf & Msg
End Sub
Sub NewName(Index As Integer, NName As String)
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM users WHERE Username = '" & Clients(Index).User & "';", db, adOpenKeyset, adLockOptimistic
rs.MoveFirst
If Err <> 0 Then Exit Sub
'rs.edit
rs!DisplayName = NName
rs.Update
SendNewNameToLists Index, NName
End Sub
