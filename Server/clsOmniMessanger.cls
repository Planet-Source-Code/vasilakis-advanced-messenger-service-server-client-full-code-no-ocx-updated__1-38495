VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msgServer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function DisplayName(who) As String
On Error Resume Next
Dim rWho As String
rWho = who
DisplayName = frmMain.GetUserDisplayName(rWho)
End Function

Public Function GetCountUsers() As Long
On Error Resume Next
GetCountUsers = frmMain.ClientsCount
End Function




Public Sub sendmsg(ToWho, FromWho, whatText)
On Error Resume Next
Dim rIndex As Integer
Dim tWho As String
Dim fWho As String
Dim wText As String
tWho = ToWho
fWho = FromWho
wText = whatText
rIndex = frmMain.GetIndexFromUsername(tWho)
frmMain.SendData rIndex, "incomingmsg;" & fWho & vbLf & " [Web Message] " & wText
End Sub

Public Function UserOnline(who) As Boolean
On Error Resume Next
Dim rWho As String
rWho = who
UserOnline = frmMain.IsOnline(rWho)
End Function


Public Function UserStatus(who) As Boolean
On Error Resume Next
Dim rWho As String
rWho = who
UserStatus = Clients(frmMain.GetIndexFromUsername(rWho)).Status
End Function
