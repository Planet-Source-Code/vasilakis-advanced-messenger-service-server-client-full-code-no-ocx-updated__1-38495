Attribute VB_Name = "modTray"

      Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallbackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

      Global Const NIM_ADD = &H0
      Global Const NIM_MODIFY = &H1
      Global Const NIM_DELETE = &H2

      Global Const WM_MOUSEMOVE = &H200

      Global Const NIF_MESSAGE = &H1
      Global Const NIF_ICON = &H2
      Global Const NIF_TIP = &H4

      Global Const WM_LBUTTONDBLCLK = &H203
      Global Const WM_LBUTTONDOWN = &H201
      Global Const WM_LBUTTONUP = &H202

      Global Const WM_RBUTTONDBLCLK = &H206
      Global Const WM_RBUTTONDOWN = &H204
      Global Const WM_RBUTTONUP = &H205

      Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Global nid As NOTIFYICONDATA








Global Const ABM_GETTASKBARPOS = &H5&

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Global Notify As NOTIFYICONDATA
Global BarData As APPBARDATA
Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long ' message specific
    End Type


Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long



Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hWnd = Form1.hWnd
    Notify.uId = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub



Sub delIcon(IconID As Long)
    Dim Result As Long
    Notify.uId = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)
End Sub







Sub modIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hWnd = Form1.hWnd
    Notify.uId = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)
End Sub



