VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFileTransfer 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin MSComctlLib.ImageList imgReceive 
      Left            =   3120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":080C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":0966
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":0C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileTransfer.frx":0D74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsockF 
      Left            =   2880
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   2280
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pFull 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label lblP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   195
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   5085
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "files remaining."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblSyn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stored 0bytes of 0bytes."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblFull 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0% completed."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PubErrors As Long
Public rLastTransfered
Public bytesPerSec As Long
Public rbytesPerSec As Long
Public closeCount
Public rHost As String
Public rFile As String
Public rStoreAs As String
Public fr
                Public rTemp
                Public rREC
Public AllFilesSize
Public TransferSize
Public Transfered

Public IP
Public port
Public Application
Public AppVersion
Public AppPath
Public NotifyOnly
Public UpdateToVersion
Public IconObjectReceive As Object
Public CRC As String
Public UnloadMe As Boolean
Function Confirm() As Boolean
Dim rYN
rYN = MsgBox("You have an incoming file transfer request from " & frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2)) & ". Would you like to accept it?", vbYesNo + vbQuestion, "File Transfer")
If rYN = vbNo Then Confirm = False Else Confirm = True
End Function

Function GetPath(file As String)
Dim iFile As String
Dim rTemp, fTemp As String
iFile = file
Do
    rTemp = InStr(iFile, "\")
    iFile = Right(iFile, Len(iFile) - rTemp)
    If rTemp = 0 Then Exit Do
    fTemp = fTemp + rTemp
Loop
GetPath = Left(file, fTemp)

End Function

Sub SendData(Text)
On Error Resume Next
wsock.SendData Text
DoEvents
End Sub






Private Sub cmdCancel_Click()
UnloadMe = True
wsock.Close
Unload Me
End Sub

Private Sub Form_Load()
UnloadMe = False
Show
AppPath = "c:\"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim Msg As Long
          Dim sFilter As String
          If ScaleMode <> 3 Then Msg = X / Screen.TwipsPerPixelX Else: Msg = X
          Select Case Msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK
                Show
                Visible = True
                WindowState = 0
             Case WM_RBUTTONDOWN
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Sub

Private Sub Form_Resize()
'Hide
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If UnloadMe Then
    delIcon Me.IconObjectReceive.Handle
Else
    Cancel = True
    Hide
End If
End Sub

Private Sub wsock_Close()
On Error Resume Next
Dim ws As Winsock
'
If lblNum.Caption = "1" Then
                Close Val(wsockF.Tag)
                UnloadMe = True
                Unload Me
                Exit Sub
End If

On Error Resume Next
Err = 0
wsock.Close
wsockF.Close
wsock.Tag = ""
Close Val(wsockF.Tag)
UnloadMe = True
Unload Me
End Sub

Private Sub wsock_DataArrival(ByVal bytesTotal As Long)
Randomize Timer
On Error Resume Next
Dim i
Dim rFileNameOnly As String
Dim m_CRC As New clsCRC
Dim rCRC, rCRC2 As String
Dim rPath As String, rFilePath  As String
Dim iCount
Dim iTemp, rTemp, rTempV, rTmp
Dim FFile As Double
Dim rPort As Double
Dim vtData As String
Dim curPOS As Single
Dim ws As Winsock
Dim FileCont As String
Dim intRetVal As Integer
Dim COMM As String, MESSAGE As String
Dim gW As Long
Dim gWd As Long
Dim wFName As String
Dim wName As String
Dim mApp As String
If wsock.State <> sckConnected Then Exit Sub
Randomize Timer
curPOS = 0
wsock.GetData vtData
wsock.Tag = wsock.Tag & vtData
GoTo CheckMSG
cmd:
curPOS = 0

            If COMM = "files" Then
                lblNum.Caption = Trim(MESSAGE + 1)
                
            ElseIf COMM = "size" Then
                TransferSize = CLng(MESSAGE)
                DoEvents
            ElseIf COMM = "fullsize" Then
                lblStatus.Caption = "Initializing..."
                AllFilesSize = CLng(MESSAGE)
                pFull.Max = AllFilesSize
                pFull.Value = pFull.Min
                DoEvents
                SendData "getfiles" & vbCrLf
            ElseIf COMM = "transfer" Then
                CRC = GetPiece(MESSAGE, vbLf, 2)
                lblStatus.Caption = GetPiece(MESSAGE, vbLf, 1)
                pBar.Max = TransferSize
                pBar.Value = pBar.Min
                lblNum.Caption = Trim(Str$(Val(lblNum.Caption) - 1))
                lblNum.Visible = True
                Me.lblFiles.Visible = True
                If Val(lblNum.Caption) = 1 Then
                    lblNum.Visible = False
                    lblFiles.Caption = "1 file remaining..."
                    lblFiles.Left = pBar.Left
                End If
                If Val(lblNum.Caption) <= 0 Then lblNum.Caption = "0"
                If GetPiece(MESSAGE, " ", 1) = "" Then
                    SendData "error;invalid_filename" & vbCrLf
                    Close Val(wsockF.Tag)
                    UnloadMe = True
                    Unload Me
                    Exit Sub
                End If
                rPath = AppPath
                If Mid(rPath, Len(rPath)) <> "\" Then
                    rPath = rPath & "\"
                End If
                rPath = rPath & "Received\"
                MkDir rPath
                rFilePath = rPath & GetPiece(MESSAGE, vbLf, 1)
                Err = 0
                Close Val(wsockF.Tag)
                Err = 0
                rCRC = GetSetting("vasilakis Messenger", "ResumeFiles", lblStatus.Caption, "")
                FFile = FreeFile
                Err = 0
                Open rFilePath For Input As FFile
                If Err = 0 Then
                    Close FFile
                    If rCRC = "" Or rCRC <> CRC Then
                        rFileNameOnly = Left(GetPiece(MESSAGE, vbLf, 1), Len(GetPiece(MESSAGE, vbLf, 1)) - Len(GetExt(GetPiece(MESSAGE, vbLf, 1))))
                        Err = 0
RetryRandomName:
    i = i + 1
                        rFilePath = rPath & rFileNameOnly & "[" & i & "]" & "." & GetExt(GetPiece(MESSAGE, vbLf, 1))
                        FFile = FreeFile
                        Open rFilePath For Input As FFile
                            If Err = 0 Then Close FFile: GoTo RetryRandomName
                        Close FFile
                    End If
                End If
                Close FFile
                Err = 0
                FFile = FreeFile
                Open rFilePath For Binary Access Write As FFile
                
                If Err = 0 Then
                    wsockF.Tag = Str$(FFile)
RandomPort:
                    For rPort = 9812 To 9822
                        wsockF.Close
                        wsockF.RemotePort = 0
                        wsockF.LocalPort = rPort
                        Err = 0
                        wsockF.listen
                        If Err <> 0 Then
                            wsockF.Close
                        Else
                            Exit For
                        End If
                    Next rPort
                    If Err <> 0 Or wsockF.State <> sckListening Then wsockF.Close: GoTo RandomPort
                    SendData "transfer;" & rPort & vbCrLf
                Else
                    Close FFile
                    SendData "error;invalid_filename" & vbCrLf
                    Close Val(wsockF.Tag)
                    UnloadMe = True
                    Unload Me
                End If
                
            ElseIf COMM = "readytostart" Then
                rCRC = GetSetting("vasilakis Messenger", "ResumeFiles", lblStatus.Caption, "")
                If LOF(Val(wsockF.Tag)) > 0 And rCRC <> "" Then
                    If rCRC = CRC Then
                        rCRC2 = Hex(m_CRC.CalculateFile(lblStatus.Caption))
                        If LOF(Val(wsockF.Tag)) >= TransferSize And rCRC2 = CRC Then
                            If Val(lblNum.Caption) = 1 Then
                                wsock_Close
                                wsockF.Close
                                Close Val(wsockF.Tag)
                                UnloadMe = True
                                Unload Me
                                Exit Sub
                            End If
                            pFull.Value = pFull.Value + LOF(Val(wsockF.Tag))
                            SendData "getfiles" & vbCrLf
                            Exit Sub
                        End If
                        If LOF(Val(wsockF.Tag)) < TransferSize Then
                            Transfered = LOF(Val(wsockF.Tag))
                            Seek Val(wsockF.Tag), LOF(Val(wsockF.Tag))
                            SendData "seekfile;" & LOF(Val(wsockF.Tag)) + 1 & vbCrLf
                            pFull.Value = pFull.Value + LOF(Val(wsockF.Tag))
                        End If
                    End If
                    SaveSetting "vasilakis Messenger", "ResumeFiles", lblStatus.Caption, CRC
                    SendData "starttransfer" & vbCrLf
                Else
                    SaveSetting "vasilakis Messenger", "ResumeFiles", lblStatus.Caption, CRC
                    SendData "starttransfer" & vbCrLf
                End If
            ElseIf COMM = "logout" Then
                SendData "logout" & vbCrLf
                wsockF.Close
                wsock_Close
            ElseIf COMM = "ping" Then
                'do nothing, its a ping to see if socket is still alive.
        End If

NoCom:
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
DoEvents
End Sub

Private Sub wsockF_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
If wsock.State <> sckConnected Then Exit Sub
If wsockF.RemoteHostIP = wsock.RemoteHostIP Then
DoEvents
DoEvents
    Transfered = 0
    wsockF.Close
    wsockF.Accept requestID
End If
End Sub

Private Sub wsockF_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim filedata As String
Dim rProc As String
Dim rProc2 As Integer
Dim rName As String
filedata = ""
If wsockF.State <> sckConnected Then Exit Sub
wsockF.GetData filedata
Transfered = Transfered + bytesTotal

pFull.Value = pFull.Value + bytesTotal
rProc = CInt(100 / pFull.Max * pFull.Value)
rProc2 = CInt(100 / pBar.Max * pBar.Value)
lblFull.Caption = rProc & "% completed."
lblSyn.Caption = "Stored " & KB(Int(pFull.Value)) & " of " & KB(Int(pFull.Max)) & "."
                
                If lblP.Caption <> rProc2 & "%" Then
                    If frmMain.wsock.State = sckConnected Then
                        rName = frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
                    Else
                        rName = GetPiece(Tag, vbLf, 2)
                    End If
                    lblP.Caption = rProc2 & "%"
                    If rProc2 < 10 Then
                        Me.Icon = imgReceive.ListImages(1).Picture
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 10 And rProc2 < 20 Then
                        Me.Icon = imgReceive.ListImages(2).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 20 And rProc2 < 30 Then
                        Me.Icon = imgReceive.ListImages(3).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 30 And rProc2 < 40 Then
                        Me.Icon = imgReceive.ListImages(4).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 40 And rProc2 < 50 Then
                        Me.Icon = imgReceive.ListImages(5).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 50 And rProc2 < 60 Then
                        Me.Icon = imgReceive.ListImages(6).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 60 And rProc2 < 70 Then
                        Me.Icon = imgReceive.ListImages(7).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 70 And rProc2 < 80 Then
                        Me.Icon = imgReceive.ListImages(8).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 80 And rProc2 < 90 Then
                        Me.Icon = imgReceive.ListImages(9).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    ElseIf rProc2 >= 90 And rProc2 < 100 Then
                        Me.Icon = imgReceive.ListImages(10).Picture
                        
                        modIcon Me, IconObjectReceive.Handle, Me.Icon, "Receiveing File - " & rName & "..." & " (" & rProc2 & "%)"
                    End If
                    DoEvents
                End If

pBar.Value = Transfered
Put Val(wsockF.Tag), , filedata
If Transfered >= TransferSize And wsockF.State = sckConnected Then
    DeleteSetting "vasilakis Messenger", "ResumeFiles", lblStatus.Caption
    If Val(lblNum.Caption) = 1 Then
        wsock_Close
        Exit Sub
    End If
    SendData "getfiles" & vbCrLf
    DoEvents
End If

End Sub


Function KB(bytes) As String
On Error Resume Next
Dim s, r
If bytes > 1023 And bytes < 1000000 Then
    KB = bytes / 1000
    s = "KB"
ElseIf bytes > 1000000 Then
    KB = bytes / 1000000
    s = "MB"
ElseIf bytes < 1023 Then
    KB = bytes
    s = "bytes"
End If
r = InStr(KB, ",")
If r <> 0 Then
    KB = Left(KB, r + 1)
End If
KB = KB & s
End Function

