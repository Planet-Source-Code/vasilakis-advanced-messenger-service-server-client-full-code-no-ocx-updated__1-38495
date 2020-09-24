VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFileSend 
   BackColor       =   &H00D3D3C3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgSend 
      Left            =   2040
      Top             =   240
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
            Picture         =   "frmFileSend.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":080C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":0966
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":0C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileSend.frx":0D74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   3960
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrUser 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   1800
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8746
   End
   Begin MSWinsockLib.Winsock wsockF 
      Left            =   2280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog comDLG 
      Left            =   2760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      MaxFileSize     =   10000
   End
   Begin VB.Label lblP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "frmFileSend.frx":0ECE
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose the file(s) you want to send."
      Height          =   675
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3930
   End
End
Attribute VB_Name = "frmFileSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Transfered
Public curFile
Public TransferSize
Public Filename
Public rREC
Public iCount
Public cur_FileBuf
Public CRC
Public SentSize

'Public IconObjectSend As Object

Sub ChooseFiles()
    On Error Resume Next
    Dim rFFile
    Dim vFiles As Variant
    Dim lFile As Long
    comDLG.Filename = ""
    comDLG.CancelError = True
    comDLG.DialogTitle = "Select File(s)..."
    comDLG.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    comDLG.Filter = "All files (*.*)|*.*"
    Err = 0
    comDLG.ShowOpen
    If Err <> 0 Then Exit Sub
    vFiles = Split(comDLG.Filename, Chr(0))
    If UBound(vFiles) = 0 Then
            rFFile = FreeFile
            Err = 0
            Open comDLG.Filename For Input As rFFile
            If Err = 0 And LOF(rFFile) > 0 Then
                Close rFFile
                lstFiles.AddItem comDLG.Filename
            Else
                MsgBox "Couldn't add file '" & vFiles(lFile) & "' because access is denied or it has a zero length size.", vbCritical, "Oops!"
            End If
    Else
        For lFile = 1 To UBound(vFiles)
            rFFile = FreeFile
            Err = 0
            Open vFiles(0) + "\" & vFiles(lFile) For Input As rFFile
            If Err = 0 And LOF(rFFile) > 0 Then
                Close rFFile
                lstFiles.AddItem vFiles(0) + "\" & vFiles(lFile)
            Else
                MsgBox "Couldn't add file '" & vFiles(lFile) & "' because access is denied or it has a zero length size.", vbCritical, "Oops!"
            End If
        Next
    End If
    Exit Sub
errrhandler:
    Unload Me


End Sub

Sub Test()
'Dim rProc As String
 '               If lblP.Caption <> rProc & "%" Then
  '                  lblP.Caption = rProc & "%"
   '                 rName = frmMain.GetNameFromEmail(GetPiece(Tag, vbLf, 2))
    '                If rProc < 10 Then
     '                   Me.Icon = imgSend.ListImages(1).Picture
      '                  modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
       '             ElseIf rProc >= 10 And rProc < 20 Then
        '                Me.Icon = imgSend.ListImages(2).Picture
         '               modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
          '          ElseIf rProc >= 20 And rProc < 30 Then
           '             Me.Icon = imgSend.ListImages(3).Picture
            '            modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
             '       ElseIf rProc >= 30 And rProc < 40 Then
              '          Me.Icon = imgSend.ListImages(4).Picture
               '         modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
                '    ElseIf rProc >= 40 And rProc < 50 Then
                 '       Me.Icon = imgSend.ListImages(5).Picture
                  '      modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
                   ' ElseIf rProc >= 50 And rProc < 60 Then
'                        Me.Icon = imgSend.ListImages(6).Picture
 '                       modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
  '                  ElseIf rProc >= 60 And rProc < 70 Then
   '                     Me.Icon = imgSend.ListImages(7).Picture
    '                    modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
     '               ElseIf rProc >= 70 And rProc < 80 Then
      '                  Me.Icon = imgSend.ListImages(8).Picture
       '                 modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
        '            ElseIf rProc >= 80 And rProc < 90 Then
         '               Me.Icon = imgSend.ListImages(9).Picture
          '              modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
           '         ElseIf rProc >= 90 And rProc < 100 Then
            '            Me.Icon = imgSend.ListImages(10).Picture
             '           modIcon Me, IconObjectSend.Handle, Me.Icon, "Sending File(s) - " & rName & "..." & " (" & rProc & "%)"
              '      End If
               ' End If

End Sub

Private Sub Form_Activate()
SetWindowPos hWnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Load()
cur_FileBuf = 5000
Visible = False
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


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tmrUser.Enabled = False
wsock.Close
wsockF.Close
DoEvents
DoEvents
Refresh
Close Val(wsockF.Tag)
End Sub



Private Sub lblStatus_Click()
'frmMain.CreateNewFileReceive "test", "194.219.197.24", wsock.LocalPort
End Sub


Private Sub tmrUser_Timer()
On Error Resume Next
Dim fr
Dim FileCont As String
Dim rProc
Dim rName As String
fr = Val(wsockF.Tag)
If Err <> 0 Then Exit Sub
                If wsockF.State <> sckConnected Then
                    tmrUser.Enabled = False
                    Exit Sub
                End If
                If LOF(fr) - Transfered < cur_FileBuf Then
                    FileCont = Input(LOF(fr) - Transfered, fr)
                    Transfered = Transfered + (LOF(fr) - Transfered)
                    DoEvents
                Else
                    FileCont = Input(cur_FileBuf, fr)
                    Transfered = Transfered + cur_FileBuf
                    DoEvents
                End If
                wsockF.SendData FileCont
                If Transfered >= LOF(fr) Then
                    tmrUser.Enabled = False
                End If
                If wsockF.State <> sckConnected Then
                    tmrUser.Enabled = False
                End If
End Sub



Private Sub wsock_Close()
Unload Me
End Sub

Private Sub wsock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next

wsock.Close
wsock.Accept requestID
lblStatus.Caption = "Connected. Setting up transfer..."
End Sub


Private Sub wsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim m_CRC As New clsCRC
Dim CRC As String
Dim FileCont As String
Dim intRetVal As Integer
Dim gW As Long
Dim gWd As Long
Dim vtData As String
Dim curPOS As Single
Dim wFName As String
Dim wName As String
Dim ws As Winsock
Dim mApp As String
Dim COMM As String, MESSAGE As String
Dim rTemp, iTemp, rTempV, rTmp
Dim rFullSize, i, rFFile, fr, rSize
Dim rFIleToTransfer  As String, rFileTransfer  As String

Randomize Timer
curPOS = 0
wsock.GetData vtData
wsock.Tag = wsock.Tag & vtData

GoTo CheckMSG
cmd:
curPOS = 0
        If COMM = "start" Then
            curFile = -1
            If lstFiles.ListCount = 0 Then
                SendData "nofiles" & vbCrLf
            End If
            'Authenticated, Setting up Transfer...
            SendData "files;" & Trim(Str$(lstFiles.ListCount)) & vbCrLf
            rFullSize = 0
            For i = 0 To lstFiles.ListCount - 1
                rFFile = FreeFile
                Open lstFiles.List(i) For Input As rFFile
                    rFullSize = rFullSize + LOF(rFFile)
                Close rFFile
            Next i
            SendData "fullsize;" & rFullSize & vbCrLf
        ElseIf COMM = "getfiles" Then

            curFile = curFile + 1
            Err = 0
            Transfered = 0
            TransferSize = 0
            rFileTransfer = lstFiles.List(curFile)
            rFIleToTransfer = Right(rFileTransfer, Len(rFileTransfer) - Len(GetPath(rFileTransfer)))
            Close Val(wsockF.Tag)
            fr = FreeFile
            Filename = rFIleToTransfer
            Open rFileTransfer For Binary Access Read As fr
            If Err <> 0 Then
                wsock_Close
                Exit Sub
            End If
            rSize = LOF(fr)
            Close fr
            DoEvents
            lblStatus.Caption = "Calculating CRC..."
            DoEvents
            CRC = Hex(m_CRC.CalculateFile(rFileTransfer))
            DoEvents
            lblStatus.Caption = "Requesting Upload..."
            SendData "size;" & rSize & vbCrLf
            DoEvents
            SendData "transfer;" & rFIleToTransfer & vbLf & CRC & vbCrLf
            DoEvents
        ElseIf COMM = "pause" Then
            If tmrUser.Enabled = True Then tmrUser.Enabled = False
        ElseIf COMM = "resume" Then
            If tmrUser.Enabled = False Then tmrUser.Enabled = True
        ElseIf COMM = "transfer" Then
            Err = 0
            SentSize = 0
            lblStatus.Caption = "Connecting..."
            wsockF.Close
            wsockF.Connect wsock.RemoteHostIP, Val(MESSAGE)

            Do
                DoEvents
            Loop Until wsockF.State = sckConnected Or wsockF.State = sckError
            tmrUser.Enabled = False
            If wsockF.State = sckConnected Then
                Err = 0
                
                fr = FreeFile
                wsockF.Tag = fr
                Open lstFiles.List(curFile) For Binary Access Read Lock Read As fr
                    If Err <> 0 Then Close fr: Unload Me: Exit Sub
                    DoEvents
                    Transfered = 0
                    TransferSize = LOF(fr)
             
' *** UPLOAD ROUTINE

             FileCont = ""

                DoEvents
                SendData "readytostart" & vbCrLf
                lblStatus.Caption = Filename
            Else
                Unload Me
            End If
        ElseIf COMM = "seekfile" Then
            Transfered = Val(MESSAGE) - 2
            SentSize = Transfered
            Seek Val(wsockF.Tag), Val(MESSAGE) - 1
        ElseIf COMM = "starttransfer" Then
            tmrUser.Interval = 10
            tmrUser.Enabled = True
        ElseIf COMM = "ack" Then
            
            'Acknowledged Packets
            'pAck.Value = Val(message)
            'If pAck.Value = pAck.Max Then
                    'DoEvents
                    'wsockF(Index).Tag = ""
                    'SendData Index, "logout" & vbCrLf
                    'DoEvents
            'End If
                
        
        ElseIf COMM = "logout" Then
            
                    wsock_Close
                    
        
        ElseIf COMM = "error" Then

            'Error Received, Check it and retry...
                    Unload Me
                    Exit Sub
                
        
        Else

            'Lathos Entolh
        
        End If
'    End If
    
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


Function GetPath(file)
On Error Resume Next
Dim iFile, rTemp, fTemp
iFile = file
Do
    rTemp = InStr(iFile, "\")
    iFile = Right(iFile, Len(iFile) - rTemp)
    If rTemp = 0 Then Exit Do
    fTemp = fTemp + rTemp
Loop
GetPath = Left(file, fTemp)
Err = 0
End Function





Sub SendData(Text)
On Error Resume Next
wsock.SendData Text
DoEvents
End Sub


Private Sub wsockF_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Dim fr
Dim FileCont As String
Dim rProc
Dim rName As String
fr = Val(wsockF.Tag)
SentSize = SentSize + bytesSent
                If SentSize <> 0 Then
                        rProc = CInt(100 / LOF(fr) * SentSize)
                End If
                lblP.Caption = "Sending File " & rProc & "%"
                Caption = rProc & "% Completed."

End Sub


