VERSION 5.00
Begin VB.Form frmBar 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "AppBar Demo"
   ClientHeight    =   645
   ClientLeft      =   3000
   ClientTop       =   1980
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SentType As Boolean
Public Focused As Boolean
Public AppBar As New TAppBar
Private Sub Form_Load()
  
  AppBar.Extends Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
AppBar.Detach
End Sub












