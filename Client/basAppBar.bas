Attribute VB_Name = "basAppBar"
Option Explicit

Public jPath As String
Public jData As String

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallBackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const WM_USER = &H400
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SM_CYSCREEN = 1
Private Const SM_CXSCREEN = 0
Private Const ABM_NEW = &H0&
Private Const ABM_REMOVE = &H1&
Private Const ABM_QUERYPOS = &H2&
Private Const ABM_SETPOS = &H3&
Private Const ABM_GETSTATE = &H4&
Private Const ABM_GETTASKBARPOS = &H5&
Private Const ABM_ACTIVATE = &H6&          'lParam == TRUE/FALSE means activate/deactivate
Private Const ABM_GETAUTOHIDEBAR = &H7&
Private Const ABM_SETAUTOHIDEBAR = &H8&
Private Const ABE_LEFT = 0
Private Const ABE_TOP = 1
Private Const ABE_RIGHT = 2
Private Const ABE_BOTTOM = 3
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90
Private Const nTwipsPerInch = 1440
Private Const GWL_STYLE = (-16)

Public Enum jPosition
    jBottom = ABE_BOTTOM
    jtop = ABE_TOP
    jLeft = ABE_LEFT
    jRight = ABE_RIGHT
End Enum

Private jABD As APPBARDATA

Public Function ConvertTwipsToPixels(nTwips As Long, nDirection As Long) As Integer
    Dim hdc As Long
    Dim nPixelsPerInch As Long
       
    hdc = GetDC(0)
    If (nDirection = 0) Then       'Horizontal
        nPixelsPerInch = GetDeviceCaps(hdc, WU_LOGPIXELSX)
    Else                            'Vertical
        nPixelsPerInch = GetDeviceCaps(hdc, WU_LOGPIXELSY)
    End If
    
    hdc = ReleaseDC(0, hdc)
    ConvertTwipsToPixels = (nTwips / nTwipsPerInch) * nPixelsPerInch
End Function

Public Sub CreateAppBar(jForm As Form, jPos As jPosition)
    With jABD
        .cbSize = Len(jABD)
        .hwnd = jForm.hwnd
        .uCallBackMessage = WM_USER + 100
    End With
    Call SHAppBarMessage(ABM_NEW, jABD)
    
    Select Case jPos
        Case jBottom
            jABD.uEdge = ABE_BOTTOM
        Case jtop
            jABD.uEdge = ABE_TOP
    End Select
    
    Call SetRect(jABD.rc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
    Call SHAppBarMessage(ABM_QUERYPOS, jABD)
    
    Select Case jPos
        Case jBottom
            jABD.rc.Top = jABD.rc.Bottom - ConvertTwipsToPixels(jForm.Height, 1)
        Case jtop
            jABD.rc.Bottom = jABD.rc.Top + ConvertTwipsToPixels(jForm.Height, 1)
    End Select
    
    Call SHAppBarMessage(ABM_SETPOS, jABD)

    Select Case jPos
        Case jBottom
            Call SetWindowPos(jABD.hwnd, 0, jABD.rc.Left, jABD.rc.Top, jABD.rc.Right - jABD.rc.Left, jABD.rc.Bottom - jABD.rc.Top, SWP_NOZORDER Or SWP_NOACTIVATE)
        Case jtop
            Call SetWindowPos(jABD.hwnd, 0, jABD.rc.Left, jABD.rc.Top, jABD.rc.Right - jABD.rc.Left, jABD.rc.Bottom - jABD.rc.Top, SWP_NOZORDER Or SWP_NOACTIVATE)
    End Select
End Sub

Public Sub DestroyAppBar()
     Call SHAppBarMessage(ABM_REMOVE, jABD)
End Sub

Public Sub AppBarActivateMsg()
    Call SHAppBarMessage(ABM_ACTIVATE, jABD)
End Sub
