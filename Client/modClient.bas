Attribute VB_Name = "modClient"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW           As Integer = 5

Public ChatForm() As New frmChat
Public ChatsIndex As Long

Public DCChatForm() As New frmDCC
Public DCChatsIndex As Long

Public AuthorizeUser() As New frmAuth
Public AuthsIndex As Long

Public FileTransfer() As New frmFileTransfer
Public FilesIndex As Long

Public FileSend() As New frmFileSend
Public FilesSendIndex As Long

Public PopupWindow() As New frmPopup
Public Popups As Long

Global Const bold = 2
Global Const underline = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1


Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String

Global strFont As String
Global strFontSize As Integer
Global lngBackColor As Long
Global lngForeColor As Long
Global lngLeftColor As Long
Global lngRightColor As Long
Global strFontName As String
Global intFontSize As Integer


Public Const ABM_NEW = &H0
Public Const ABM_REMOVE = &H1
Public Const ABM_QUERYPOS = &H2
Public Const ABM_SETPOS = &H3
Public Const ABM_GETSTATE = &H4
Public Const ABM_GETTASKBARPOS = &H5
Public Const ABM_ACTIVATE = &H6
Public Const ABM_GETAUTOHIDEBAR = &H7
Public Const ABM_SETAUTOHIDEBAR = &H8
Public Const ABM_WINDOWPOSCHANGED = &H9
Public Const ABN_STATECHANGE = &H0
Public Const ABN_POSCHANGED = &H1
Public Const ABN_FULLSCREENAPP = &H2
Public Const ABN_WINDOWARRANGE = &H3
Public Const ABS_AUTOHIDE = &H1
Public Const ABS_ALWAYSONTOP = &H2
Public Const ABE_LEFT = 0
Public Const ABE_TOP = 1
Public Const ABE_RIGHT = 2
Public Const ABE_BOTTOM = 3
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE



Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Const SPI_GETWORKAREA = 48

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40

Public Const HKEY_LOCAL_MACHINE = &H80000002

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
Hkey As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
"RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" _
Alias "DefWindowProcA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
Public Const WM_SETHOTKEY = &H32
Public Const WM_SHOWWINDOW = &H18
Public Const HK_SHIFTA = &H141 'Shift + A
Public Const HK_SHIFTB = &H142 'Shift * B
Public Const HK_CONTROLA = &H241 'Control + A
Public Const HK_ALTZ = &H45A

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long


Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long



Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Function GetEmailProgram() As String
On Error Resume Next
rPr = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
GetEmailProgram = GetPiece2(Replace(getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\mailto\shell\open\command", ""), "%ProgramFiles%", rPr), Chr(34), 1)
End Function

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
End Function


Public Function ActiveConnection() As Boolean
Dim Hkey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
ActiveConnection = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, _
phkResult)

If ReturnCode = ERROR_SUCCESS Then
    Hkey = phkResult
    lpValueName = "Remote Connection"
    lpReserved = APINULL
    lpType = APINULL
    lpData = APINULL
    lpcbData = APINULL
    ReturnCode = RegQueryValueEx(Hkey, lpValueName, _
    lpReserved, lpType, ByVal lpData, lpcbData)
    lpcbData = Len(lpData)
    ReturnCode = RegQueryValueEx(Hkey, lpValueName, _
    lpReserved, lpType, lpData, lpcbData)
    
    If ReturnCode = ERROR_SUCCESS Then
        If lpData = 0 Then
            ActiveConnection = False
        Else
            ActiveConnection = True
        End If
    End If
                
RegCloseKey (Hkey)
End If
End Function
Function GetPiece2(from, delim, Index) As String
rParms = Split(from, delim)
rIndex = 0
For i = 0 To UBound(rParms)
    If rParms(i) <> "" Then
        rIndex = rIndex + 1
        If rIndex = Index Then
            GetPiece2 = rParms(i)
        End If
    End If
Next i
End Function

Function AnsiColor(intColNum As Integer) As Long
    Select Case intColNum
        Case 0: AnsiColor = RGB(255, 255, 255)
        Case 1: AnsiColor = RGB(0, 0, 0)
        Case 2: AnsiColor = RGB(0, 0, 127)
        Case 3: AnsiColor = RGB(0, 127, 0)
        Case 4: AnsiColor = RGB(255, 0, 0)
        Case 5: AnsiColor = RGB(127, 0, 0)
        Case 6: AnsiColor = RGB(127, 0, 127)
        Case 7: AnsiColor = RGB(255, 127, 0)
        Case 8: AnsiColor = RGB(255, 255, 0)
        Case 9: AnsiColor = RGB(0, 255, 0)
        Case 10: AnsiColor = RGB(0, 148, 144)
        Case 11: AnsiColor = RGB(0, 255, 255)
        Case 12: AnsiColor = RGB(0, 0, 255)
        Case 13: AnsiColor = RGB(255, 0, 255)
        Case 14: AnsiColor = RGB(92, 92, 92)
        Case 15: AnsiColor = RGB(184, 184, 184)
        Case Else: AnsiColor = RGB(0, 0, 0)
    End Select
End Function





Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function



Function LeftR(strData As String, intMin As Integer)
    
    On Error Resume Next
    LeftR = Left(strData, Len(strData) - intMin)
End Function



Sub Main()
Load frmSplash
frmSplash.Show
frmSplash.Refresh
DoEvents
DoEvents
Load frmMain
End Sub



Function GetExt(file As String)
iFile = file
Do
    rTemp = InStr(iFile, ".")
    iFile = Right(iFile, Len(iFile) - rTemp)
    If rTemp = 0 Then Exit Do
    fTemp = fTemp + rTemp
Loop
GetExt = "." & Right(file, Len(file) - fTemp)
End Function



Sub PutData(RTF As RichTextBox, strData As String, showLast As Boolean)
    lngBackColor = RGB(255, 255, 255)
    lngForeColor = RGB(0, 0, 0)
    lngLeftColor = &H800000
    lngRightColor = &H8000000F
    strFontName = "Tahoma"
    intFontSize = 10

    'MsgBox InStr(strData, "W")
    If 0 = InStr(strData, strBold) And _
       0 = InStr(strData, strUnderline) And _
       0 = InStr(strData, strReverse) And _
       0 = InStr(strData, strColor) Then
       RTF.SelStart = Len(RTF.Text)
       RTF.SelColor = lngForeColor
       RTF.SelBold = False
       RTF.SelUnderline = False
       RTF.SelStrikeThru = False
       RTF.SelFontName = strFontName
       RTF.SelFontSize = intFontSize
       RTF.SelText = " " & strData & vbCrLf
       Exit Sub
    End If
    
    If strData = "" Then Exit Sub
    'DoEvents
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    strData = " " & strData
    Length = Len(strData)
    i = 1
    RTF.SelStart = Len(RTF.Text)
    RTF.SelColor = lngForeColor
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    RTF.SelFontName = strFontName
    RTF.SelFontSize = intFontSize
    Do
        strChar = Mid(strData, i, 1)
        Select Case strChar
            Case strBold, Chr(15)
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelBold = Not RTF.SelBold
                i = i + 1
            Case strUnderline
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelUnderline = Not RTF.SelUnderline
                i = i + 1
            Case strReverse
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                RTF.SelStrikeThru = Not RTF.SelStrikeThru
                i = i + 1
            Case strColor
                RTF.SelStart = Len(RTF.Text)
                RTF.SelText = strBuffer
                strBuffer = ""
                i = i + 1
                If i > Length Then GoTo TheEnd
                Do Until Not ValidColorCode(strBuffer) Or i > Length
                    strBuffer = strBuffer & Mid(strData, i, 1)
                    i = i + 1
                Loop
                If ValidColorCode(strBuffer) And i > Length Then GoTo TheEnd
                strBuffer = LeftR(strBuffer, 1)
                RTF.SelStart = Len(RTF.Text)
                If strBuffer = "" Then
                    RTF.SelColor = lngForeColor
                Else
                    RTF.SelColor = AnsiColor(LeftOf(strBuffer, ","))
                End If
                i = i - 1
                strBuffer = ""
            Case Else
                strBuffer = strBuffer & strChar
                i = i + 1
        End Select
    Loop Until i > Length
    If strBuffer <> "" Then
            RTF.SelStart = Len(RTF.Text)
            RTF.SelText = strBuffer
            strBuffer = ""
    End If
TheEnd:
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelStrikeThru = False
    RTF.SelStart = Len(RTF.Text)
    RTF.SelText = vbCrLf
If showLast Then
    RTF.Parent.sbBar.Panels(2).Text = "Last: " & Hour(Now) & "." & Minute(Now) & " " & Day(Now) & "/" & Month(Now)
End If
End Sub

Function ValidColorCode(strCode As String) As Boolean
    'MsgBox strCode
    Dim c1 As Integer, c2 As Integer
    If strCode Like "" Or _
       strCode Like "#" Or _
       strCode Like "##" Or _
       strCode Like "#,#" Or _
       strCode Like "##,#" Or _
       strCode Like "#,##" Or _
       strCode Like "#," Or _
       strCode Like "##," Or _
       strCode Like "##,##" Or _
       strCode Like ",#" Or _
       strCode Like ",##" Then
        Dim strCol() As String
        strCol = Split(strCode, ",")
        'DoEvents
        If UBound(strCol) = -1 Then
            ValidColorCode = True
        ElseIf UBound(strCol) = 0 Then
            If strCol(0) = "" Then strCol(0) = 0
            If CInt(strCol(0)) >= 0 And CInt(strCol(0)) < 16 Then
                ValidColorCode = True
            Else
                ValidColorCode = False
            End If
        Else
            If strCol(0) = "" Then strCol(0) = lngForeColor
            If strCol(1) = "" Then strCol(1) = 0
            c1 = CInt(strCol(0))
            c2 = CInt(strCol(1))
            If c2 < 0 Or c2 > 16 Then
                ValidColorCode = False
            Else
                ValidColorCode = True
            End If
        End If
        ValidColorCode = True
    Else
        ValidColorCode = False
    End If
End Function




Function GetPiece(from, delim, Index) As String
    Dim temp$
    Dim Count
    Dim Where
    '
    temp$ = from & delim
    Where = InStr(temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(temp$, Where - 1)
            Exit Function
        End If
        temp$ = Right$(temp$, Len(temp$) - Where)
        Where = InStr(temp$, delim)
    'Doevents
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function

Function DecryptPassword(Number As Byte, _
    EncryptedPassword) As String
    Dim Password As String
    Dim temp As Integer
    Counter = 1
    Do Until Counter = Len(EncryptedPassword) + 1
        temp = Asc(Mid(EncryptedPassword, _
            Counter, 1)) Xor (10 - Number)
        'see if even
        If Counter Mod 2 = 0 Then
            temp = temp + Number
        Else
            temp = temp - Number
        End If
        Password = Password & Chr$(temp)
        Counter = Counter + 1
    Loop
    DecryptPassword = Password
End Function




Function EncryptPassword(Number As Byte, _
    DecryptedPassword) As String
    Dim Password As String
    Dim temp As Integer
    Counter = 1
    Do Until Counter = Len(DecryptedPassword) + 1
        temp = Asc(Mid(DecryptedPassword, Counter, 1))
        'see if even
        If Counter Mod 2 = 0 Then
            temp = temp - Number
        Else
            temp = temp + Number
        End If
        temp = temp Xor (10 - Number)
        Password = Password & Chr$(temp)
        Counter = Counter + 1
    Loop
    EncryptPassword = Password
End Function











