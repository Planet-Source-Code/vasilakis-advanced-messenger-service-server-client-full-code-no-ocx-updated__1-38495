VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banner Server"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh DB"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
   End
   Begin VB.ListBox lstBannerID 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   2880
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstBannerImage 
      Height          =   2205
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstBannerURL 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IconObject As Object

Public db As New ADODB.Connection

Public cyclebanner
Function GetURL(BannerID As String) As String
On Error Resume Next
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM banners WHERE BannerID = '" & BannerID & "';", db
rs.MoveFirst
GetURL = rs.Fields("URL")
End Function

Function GetImage(BannerID As String) As String
On Error Resume Next
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM banners WHERE BannerID = '" & BannerID & "';", db
rs.MoveFirst
GetImage = rs.Fields("ImagePath")
End Function


Sub LoadBanners()
On Error Resume Next
Dim rs As New ADODB.Recordset
rs.Open "banners", db, 2, 2
rs.MoveFirst
While Not rs.EOF
    rBannerID = rs.Fields("BannerID")
    rBannerImage = rs.Fields("ImagePath")
    rBannerurl = rs.Fields("URL")
    lstBannerID.AddItem rBannerID
    lstBannerImage.AddItem rBannerImage
    lstBannerURL.AddItem rBannerurl
    rs.MoveNext
Wend
rs.Close
End Sub

Sub SendFile(Index As Integer)
On Error Resume Next
FreeF = FreeFile
Open GetImage(wsock(Index).Tag) For Binary Access Read As #FreeF
    vtdata = Input(LOF(FreeF), #FreeF)
Close #FreeF
wsock(Index).SendData vtdata
DoEvents
End Sub

Private Sub cmdRefresh_Click()
Me.lstBannerID.Clear
Me.lstBannerImage.Clear
Me.lstBannerURL.Clear
Dim ws As Winsock
For Each ws In wsock
    If ws.Index > 0 Then
        ws.Close
        Unload ws
    End If
Next ws
delIcon IconObject.Handle
Form_Load
End Sub

Private Sub Form_Load()
On Error Resume Next
Hide
Set IconObject = Me.Icon
AddIcon Me, IconObject.Handle, IconObject, "vasilakis :messenger! [Banner Server]"

If App.PrevInstance Then End
ConnectionString = "driver={SQL Server};" & _
        "server=production;" & _
        "uid=sa;pwd=;" & _
        "database=vsag;"

db.ConnectionString = ConnectionString
db.ConnectionTimeout = 10
db.Open
LoadBanners
cyclebanner = 0
lstBannerImage.ListIndex = cyclebanner
lstBannerURL.ListIndex = cyclebanner

wsock(0).LocalPort = 8990
wsock(0).Listen
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
                    Visible = True
                    Show
             Case WM_RBUTTONDOWN
             Case WM_RBUTTONUP
               
             Case WM_RBUTTONDBLCLK
          End Select
End Sub


Private Sub Form_Resize()
Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
delIcon IconObject.Handle
DoEvents
End
End Sub


Private Sub wsock_Close(Index As Integer)
Unload wsock(Index)
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim ws As Winsock
Dim ic As Integer
ic = 0
Do
ic = ic + 1
 Err = 0
 Load wsock(ic)
 If Err = 0 Or (wsock(ic).State <> sckConnected And wsock(ic).State <> sckConnecting) Then
    wsock(ic).Tag = lstBannerID.List(lstBannerURL.ListIndex)
    Err = 0
    lstBannerURL.ListIndex = lstBannerURL.ListIndex + 1
    If Err <> 0 Then lstBannerURL.ListIndex = 0
    wsock(ic).Accept requestID
    DoEvents
    wsock(ic).SendData GetURL(wsock(ic).Tag)
    'wsock(ic).SendData "http://www.vasilakis.com/"
    DoEvents
  Exit Sub
 End If
DoEvents
Loop

End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim vtdata As String
wsock(Index).GetData vtdata

Select Case vtdata
    Case "SIZE"
        FFile = FreeFile
        Open GetImage(wsock(Index).Tag) For Input As #FFile
            wsock(Index).SendData Str$(LOF(FFile))
        Close #FFile
        DoEvents
    Case "FILE"
        
        FFile = FreeFile
        SendFile Index

        DoEvents
End Select
DoEvents
End Sub


