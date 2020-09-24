VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmAppBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dock Properties"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   420
      Width           =   5295
      Begin VB.Frame grpMainWnd 
         Caption         =   "AppBar Window"
         Height          =   1305
         Left            =   180
         TabIndex        =   46
         Top             =   180
         Width           =   1665
         Begin VB.CheckBox chkAutohide 
            Caption         =   "Auto &Hide"
            Height          =   330
            Left            =   135
            TabIndex        =   2
            Top             =   750
            Width           =   1170
         End
         Begin VB.CheckBox chkAlwaysOnTop 
            Caption         =   "Always On &Top"
            Height          =   330
            Left            =   135
            TabIndex        =   1
            Top             =   330
            Width           =   1380
         End
      End
      Begin VB.Frame grpTaskEntry 
         Caption         =   "Taskbar Entry"
         Height          =   1305
         Left            =   1980
         TabIndex        =   47
         Top             =   180
         Width           =   3090
         Begin VB.OptionButton optDependEntry 
            Caption         =   "Only if &floating (abtFloatDependent)"
            Height          =   285
            Left            =   135
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   945
            Width           =   2835
         End
         Begin VB.OptionButton optHideEntry 
            Caption         =   "Hi&de (abtHide)"
            Height          =   285
            Left            =   135
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   2835
         End
         Begin VB.OptionButton optShowEntry 
            Caption         =   "&Show (abtShow)"
            Height          =   285
            Left            =   135
            TabIndex        =   3
            Top             =   255
            Width           =   2835
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   420
      Width           =   5295
      Begin VB.Frame grpFlags 
         Caption         =   "Flags"
         Height          =   2220
         Left            =   2310
         TabIndex        =   49
         Top             =   180
         Width           =   2760
         Begin VB.CheckBox chkAllowBottom 
            Caption         =   "Allow Botto&m (abfAllowBottom)"
            Height          =   210
            Left            =   150
            TabIndex        =   15
            Top             =   1860
            Width           =   2490
         End
         Begin VB.CheckBox chkAllowRight 
            Caption         =   "Allow Rig&ht (abfAllowRight)"
            Height          =   210
            Left            =   150
            TabIndex        =   14
            Top             =   1485
            Width           =   2490
         End
         Begin VB.CheckBox chkAllowTop 
            Caption         =   "Allow To&p (abeAllowTop)"
            Height          =   210
            Left            =   150
            TabIndex        =   13
            Top             =   1110
            Width           =   2490
         End
         Begin VB.CheckBox chkAllowLeft 
            Caption         =   "Allow L&eft (abfAllowLeft)"
            Height          =   210
            Left            =   150
            TabIndex        =   12
            Top             =   735
            Width           =   2490
         End
         Begin VB.CheckBox chkAllowFloat 
            Caption         =   "Allow Fl&oat (abfAllowFloat)"
            Height          =   210
            Left            =   150
            TabIndex        =   11
            Top             =   360
            Width           =   2490
         End
      End
      Begin VB.Frame grpEdge 
         Caption         =   "Edge"
         Height          =   2220
         Left            =   180
         TabIndex        =   48
         Top             =   180
         Width           =   2025
         Begin VB.OptionButton optEdgeBottom 
            Caption         =   "&Bottom (abeBottom)"
            Height          =   300
            Left            =   135
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1770
         End
         Begin VB.OptionButton optEdgeRight 
            Caption         =   "&Right (abeRight)"
            Height          =   300
            Left            =   135
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1770
         End
         Begin VB.OptionButton optEdgeTop 
            Caption         =   "&Top (abeTop)"
            Height          =   300
            Left            =   135
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1770
         End
         Begin VB.OptionButton optEdgeLeft 
            Caption         =   "&Left (abeLeft)"
            Height          =   300
            Left            =   135
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   675
            Width           =   1770
         End
         Begin VB.OptionButton optEdgeFloat 
            Caption         =   "&Float (abeFloat)"
            Height          =   300
            Left            =   135
            TabIndex        =   6
            Top             =   300
            Width           =   1770
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   420
      Width           =   5295
      Begin VB.Frame grpSizeInc 
         Caption         =   "Size Increments"
         Height          =   1710
         Left            =   180
         TabIndex        =   73
         Top             =   180
         Width           =   4890
         Begin VB.TextBox edtHorzSizeInc 
            Height          =   315
            Left            =   2280
            TabIndex        =   16
            Top             =   315
            Width           =   480
         End
         Begin VB.TextBox edtVertSizeInc 
            Height          =   315
            Left            =   2280
            TabIndex        =   17
            Top             =   810
            Width           =   480
         End
         Begin ComCtl2.UpDown updVertSizeInc 
            Height          =   360
            Left            =   2760
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtVertSizeInc"
            BuddyDispid     =   196635
            OrigLeft        =   2760
            OrigTop         =   780
            OrigRight       =   3000
            OrigBottom      =   1140
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updHorzSizeInc 
            Height          =   360
            Left            =   2760
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtHorzSizeInc"
            BuddyDispid     =   196634
            OrigLeft        =   2760
            OrigTop         =   285
            OrigRight       =   3000
            OrigBottom      =   645
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblHorzSizeInc 
            AutoSize        =   -1  'True
            Caption         =   "&Horizontal Size Increment"
            Height          =   195
            Left            =   270
            TabIndex        =   78
            Top             =   360
            Width           =   1800
         End
         Begin VB.Label lblVertSizeInc 
            AutoSize        =   -1  'True
            Caption         =   "&Vertical Size Increment"
            Height          =   195
            Left            =   270
            TabIndex        =   77
            Top             =   840
            Width           =   1620
         End
         Begin VB.Label lblZeroIncHint 
            AutoSize        =   -1  'True
            Caption         =   "Hint: Zero increments prevent resizing."
            Height          =   195
            Left            =   270
            TabIndex        =   76
            Top             =   1320
            Width           =   2715
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   3
      Left            =   120
      TabIndex        =   44
      Top             =   420
      Width           =   5295
      Begin VB.Frame grpHorzDock 
         Caption         =   "Horizontal Height"
         Height          =   2235
         Left            =   2685
         TabIndex        =   80
         Top             =   180
         Width           =   2385
         Begin VB.TextBox edtMaxHorzDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   1500
            Width           =   480
         End
         Begin VB.TextBox edtMinHorzDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   420
            Width           =   480
         End
         Begin VB.TextBox edtHorzDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            Top             =   960
            Width           =   480
         End
         Begin ComCtl2.UpDown updHorzDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   930
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtHorzDockSize"
            BuddyDispid     =   196644
            OrigLeft        =   1650
            OrigTop         =   945
            OrigRight       =   1890
            OrigBottom      =   1305
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMinHorzDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMinHorzDockSize"
            BuddyDispid     =   196643
            OrigLeft        =   1680
            OrigTop         =   405
            OrigRight       =   1920
            OrigBottom      =   765
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMaxHorzDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1470
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMaxHorzDockSize"
            BuddyDispid     =   196642
            OrigLeft        =   1650
            OrigTop         =   1485
            OrigRight       =   1890
            OrigBottom      =   1845
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMaxHorzDockSize 
            AutoSize        =   -1  'True
            Caption         =   "M&ax"
            Height          =   195
            Left            =   300
            TabIndex        =   92
            Top             =   1545
            Width           =   300
         End
         Begin VB.Label lblMinHorzDockSize 
            AutoSize        =   -1  'True
            Caption         =   "Mi&n"
            Height          =   195
            Left            =   300
            TabIndex        =   91
            Top             =   465
            Width           =   255
         End
         Begin VB.Label lblHorzDockSize 
            AutoSize        =   -1  'True
            Caption         =   "C&urrent"
            Height          =   195
            Left            =   300
            TabIndex        =   90
            Top             =   1005
            Width           =   510
         End
      End
      Begin VB.Frame grpVertDock 
         Caption         =   "Vertical Width"
         Height          =   2235
         Left            =   180
         TabIndex        =   79
         Top             =   180
         Width           =   2385
         Begin VB.TextBox edtMaxVertDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   1500
            Width           =   480
         End
         Begin VB.TextBox edtMinVertDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   420
            Width           =   480
         End
         Begin VB.TextBox edtVertDockSize 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   960
            Width           =   480
         End
         Begin ComCtl2.UpDown updVertDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   930
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtVertDockSize"
            BuddyDispid     =   196654
            OrigLeft        =   1605
            OrigTop         =   765
            OrigRight       =   1845
            OrigBottom      =   1125
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMinVertDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMinVertDockSize"
            BuddyDispid     =   196653
            OrigLeft        =   1605
            OrigTop         =   315
            OrigRight       =   1845
            OrigBottom      =   675
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMaxVertDockSize 
            Height          =   360
            Left            =   1680
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   1470
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMaxVertDockSize"
            BuddyDispid     =   196652
            OrigLeft        =   1605
            OrigTop         =   1380
            OrigRight       =   1845
            OrigBottom      =   1740
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMaxVertDockSize 
            AutoSize        =   -1  'True
            Caption         =   "Ma&x"
            Height          =   195
            Left            =   300
            TabIndex        =   89
            Top             =   1545
            Width           =   300
         End
         Begin VB.Label lblMinVertDockSize 
            AutoSize        =   -1  'True
            Caption         =   "&Min"
            Height          =   195
            Left            =   300
            TabIndex        =   88
            Top             =   465
            Width           =   255
         End
         Begin VB.Label lblVertDockSize 
            AutoSize        =   -1  'True
            Caption         =   "&Current"
            Height          =   195
            Left            =   300
            TabIndex        =   87
            Top             =   1005
            Width           =   510
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   4
      Left            =   120
      TabIndex        =   45
      Top             =   420
      Width           =   5295
      Begin VB.Frame grpMinMax 
         Caption         =   "MinMax"
         Height          =   2220
         Left            =   2550
         TabIndex        =   51
         Top             =   180
         Width           =   2520
         Begin VB.TextBox edtMinWidth 
            Height          =   315
            Left            =   1485
            TabIndex        =   28
            Top             =   300
            Width           =   480
         End
         Begin VB.TextBox edtMinHeight 
            Height          =   315
            Left            =   1485
            TabIndex        =   29
            Top             =   780
            Width           =   480
         End
         Begin VB.TextBox edtMaxWidth 
            Height          =   315
            Left            =   1485
            TabIndex        =   30
            Top             =   1275
            Width           =   480
         End
         Begin VB.TextBox edtMaxHeight 
            Height          =   315
            Left            =   1485
            TabIndex        =   31
            Top             =   1755
            Width           =   480
         End
         Begin ComCtl2.UpDown updMaxHeight 
            Height          =   360
            Left            =   1965
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1725
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMaxHeight"
            BuddyDispid     =   196665
            OrigLeft        =   3105
            OrigTop         =   1725
            OrigRight       =   3345
            OrigBottom      =   2085
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMaxWidth 
            Height          =   360
            Left            =   1965
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1245
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMaxWidth"
            BuddyDispid     =   196664
            OrigLeft        =   3375
            OrigTop         =   1275
            OrigRight       =   3615
            OrigBottom      =   1815
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMinHeight 
            Height          =   360
            Left            =   1965
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   750
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMinHeight"
            BuddyDispid     =   196663
            OrigLeft        =   3585
            OrigTop         =   750
            OrigRight       =   3825
            OrigBottom      =   1095
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updMinWidth 
            Height          =   360
            Left            =   1965
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtMinWidth"
            BuddyDispid     =   196662
            OrigLeft        =   3270
            OrigTop         =   300
            OrigRight       =   3510
            OrigBottom      =   690
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMaxHeight 
            AutoSize        =   -1  'True
            Caption         =   "Max &Height"
            Height          =   195
            Left            =   285
            TabIndex        =   59
            Top             =   1800
            Width           =   810
         End
         Begin VB.Label lblMaxWidth 
            AutoSize        =   -1  'True
            Caption         =   "Max &Width"
            Height          =   195
            Left            =   285
            TabIndex        =   58
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label lblMinHeight 
            AutoSize        =   -1  'True
            Caption         =   "Min H&eight"
            Height          =   195
            Left            =   285
            TabIndex        =   57
            Top             =   825
            Width           =   765
         End
         Begin VB.Label lblMinWidth 
            AutoSize        =   -1  'True
            Caption         =   "Min Wi&dth"
            Height          =   195
            Left            =   285
            TabIndex        =   56
            Top             =   345
            Width           =   720
         End
      End
      Begin VB.Frame grpFloatCoords 
         Caption         =   "Coords"
         Height          =   2220
         Left            =   180
         TabIndex        =   50
         Top             =   180
         Width           =   2235
         Begin VB.TextBox edtFloatLeft 
            Height          =   315
            Left            =   1215
            TabIndex        =   24
            Top             =   300
            Width           =   480
         End
         Begin VB.TextBox edtFloatTop 
            Height          =   315
            Left            =   1215
            TabIndex        =   25
            Top             =   780
            Width           =   480
         End
         Begin VB.TextBox edtFloatRight 
            Height          =   315
            Left            =   1215
            TabIndex        =   26
            Top             =   1275
            Width           =   480
         End
         Begin VB.TextBox edtFloatBottom 
            Height          =   315
            Left            =   1215
            TabIndex        =   27
            Top             =   1755
            Width           =   480
         End
         Begin ComCtl2.UpDown updFloatBottom 
            Height          =   360
            Left            =   1695
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1710
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtFloatBottom"
            BuddyDispid     =   196678
            OrigLeft        =   3105
            OrigTop         =   1725
            OrigRight       =   3345
            OrigBottom      =   2085
            Max             =   9999
            Min             =   -9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updFloatRight 
            Height          =   360
            Left            =   1695
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1245
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtFloatRight"
            BuddyDispid     =   196677
            OrigLeft        =   3375
            OrigTop         =   1275
            OrigRight       =   3615
            OrigBottom      =   1815
            Max             =   9999
            Min             =   -9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updFloatTop 
            Height          =   360
            Left            =   1695
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   750
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtFloatTop"
            BuddyDispid     =   196676
            OrigLeft        =   3585
            OrigTop         =   750
            OrigRight       =   3825
            OrigBottom      =   1095
            Max             =   9999
            Min             =   -9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown updFloatLeft 
            Height          =   360
            Left            =   1695
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            BuddyControl    =   "edtFloatLeft"
            BuddyDispid     =   196675
            OrigLeft        =   3270
            OrigTop         =   300
            OrigRight       =   3510
            OrigBottom      =   690
            Max             =   9999
            Min             =   -9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblFloatBottom 
            AutoSize        =   -1  'True
            Caption         =   "&Bottom"
            Height          =   195
            Left            =   255
            TabIndex        =   55
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label lblFloatRight 
            AutoSize        =   -1  'True
            Caption         =   "&Right"
            Height          =   195
            Left            =   255
            TabIndex        =   54
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblFloatTop 
            AutoSize        =   -1  'True
            Caption         =   "&Top"
            Height          =   195
            Left            =   255
            TabIndex        =   53
            Top             =   825
            Width           =   285
         End
         Begin VB.Label lblFloatLeft 
            AutoSize        =   -1  'True
            Caption         =   "&Left"
            Height          =   195
            Left            =   255
            TabIndex        =   52
            Top             =   345
            Width           =   270
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   5
      Left            =   120
      TabIndex        =   68
      Top             =   420
      Width           =   5295
      Begin VB.CheckBox chkSlideEffect 
         Caption         =   "&Slide Effect"
         Height          =   330
         Left            =   345
         TabIndex        =   32
         Top             =   120
         Width           =   1110
      End
      Begin VB.Frame grpSlideEffect 
         Caption         =   "                          "
         Height          =   1110
         Left            =   180
         TabIndex        =   69
         Top             =   180
         Width           =   4890
         Begin ComctlLib.Slider sldSlideTime 
            Height          =   450
            Left            =   1530
            TabIndex        =   33
            Top             =   300
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   794
            _Version        =   327682
            LargeChange     =   100
            SmallChange     =   50
            Min             =   100
            Max             =   1000
            SelStart        =   100
            TickFrequency   =   100
            Value           =   100
         End
         Begin VB.Label lblSlower 
            AutoSize        =   -1  'True
            Caption         =   "Slower"
            Height          =   195
            Left            =   4080
            TabIndex        =   72
            Top             =   765
            Width           =   480
         End
         Begin VB.Label lblFaster 
            AutoSize        =   -1  'True
            Caption         =   "Faster"
            Height          =   195
            Left            =   1515
            TabIndex        =   71
            Top             =   765
            Width           =   435
         End
         Begin VB.Label lblSlideTime 
            AutoSize        =   -1  'True
            Caption         =   "Slide &Time"
            Height          =   195
            Left            =   165
            TabIndex        =   70
            Top             =   510
            Width           =   735
         End
      End
   End
   Begin VB.Frame tabPage 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   6
      Left            =   120
      TabIndex        =   93
      Top             =   420
      Width           =   5295
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4185
         TabIndex        =   38
         Top             =   1830
         Width           =   885
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "&Load"
         Height          =   375
         Left            =   3225
         TabIndex        =   37
         Top             =   1830
         Width           =   885
      End
      Begin VB.Frame grpSettingsLocation 
         Caption         =   "Settings Location"
         Height          =   1575
         Left            =   180
         TabIndex        =   94
         Top             =   180
         Width           =   4890
         Begin VB.TextBox edtKeyName 
            Height          =   315
            Left            =   1095
            TabIndex        =   36
            Top             =   1005
            Width           =   3555
         End
         Begin VB.OptionButton optLocalMachine 
            Caption         =   "Local &Machine (HKEY_LOCAL_MACHINE)"
            Height          =   300
            Left            =   240
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   600
            Width           =   3495
         End
         Begin VB.OptionButton optCurrentUser 
            Caption         =   "Current &User (HKEY_CURRENT_USER)"
            Height          =   300
            Left            =   240
            TabIndex        =   34
            Top             =   285
            Width           =   3495
         End
         Begin VB.Label lblKeyName 
            AutoSize        =   -1  'True
            Caption         =   "Key &Name"
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   1050
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4260
      TabIndex        =   40
      Top             =   3180
      Width           =   1200
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   390
      Left            =   2985
      TabIndex        =   39
      Top             =   3180
      Width           =   1200
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   3000
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5292
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   7
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Appearance"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Position"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sizing"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Docking"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Floating"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sliding"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Registry"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAppBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Private Function CBoolInt(bVal As Boolean) As Integer
  If bVal Then
    CBoolInt = 1
  Else
    CBoolInt = 0
  End If
End Function

Private Sub InitDialog()
    
  With frmChat.AppBar
  
    ' Appearance Page
    chkAlwaysOnTop.Value = CBoolInt(.AlwaysOnTop)
    chkAutohide.Value = CBoolInt(.AutoHide)
    optShowEntry.Value = CBool(.TaskEntry = abtShow)
    optHideEntry.Value = CBool(.TaskEntry = abtHide)
    optDependEntry.Value = CBool(.TaskEntry = abtFloatDependent)

    ' Position Page
    optEdgeFloat.Value = CBool(.Edge = abeFloat)
    optEdgeLeft.Value = CBool(.Edge = abeLeft)
    optEdgeTop.Value = CBool(.Edge = abeTop)
    optEdgeRight.Value = CBool(.Edge = abeRight)
    optEdgeBottom.Value = CBool(.Edge = abeBottom)
    
    chkAllowFloat.Value = CBoolInt(.Flags And abfAllowFloat)
    chkAllowLeft.Value = CBoolInt(.Flags And abfAllowLeft)
    chkAllowTop.Value = CBoolInt(.Flags And abfAllowTop)
    chkAllowRight.Value = CBoolInt(.Flags And abfAllowRight)
    chkAllowBottom.Value = CBoolInt(.Flags And abfAllowBottom)

    ' Sizing Page
    updHorzSizeInc.Value = .HorzSizeInc
    updVertSizeInc.Value = .VertSizeInc

    ' Docking Page
    updMinHorzDockSize.Value = .MinHorzDockSize
    updMinVertDockSize.Value = .MinVertDockSize
    updHorzDockSize.Value = .HorzDockSize
    updVertDockSize.Value = .VertDockSize
    updMaxHorzDockSize.Value = .MaxHorzDockSize
    updMaxVertDockSize.Value = .MaxVertDockSize

    ' Floating Page
    updFloatLeft.Value = .FloatLeft
    updFloatTop.Value = .FloatTop
    updFloatRight.Value = .FloatRight
    updFloatBottom.Value = .FloatBottom

    updMinWidth.Value = .MinWidth
    updMinHeight.Value = .MinHeight
    updMaxWidth.Value = .MaxWidth
    updMaxHeight.Value = .MaxHeight
    
    ' Sliding Page
    chkSlideEffect.Value = CBoolInt(.SlideEffect)
    sldSlideTime.Value = .SlideTime
    sldSlideTime.Enabled = CBool(chkSlideEffect.Value)
    lblSlideTime.Enabled = CBool(chkSlideEffect.Value)
    lblFaster.Enabled = CBool(chkSlideEffect.Value)
    lblSlower.Enabled = CBool(chkSlideEffect.Value)
    
    ' Registry Page
    optCurrentUser.Value = CBool(.RootKey = HKEY_CURRENT_USER)
    optLocalMachine.Value = CBool(.RootKey = HKEY_LOCAL_MACHINE)
    edtKeyName.Text = .KeyName
  
  End With
  
End Sub

Private Sub ApplyChanges()

  With frmChat.AppBar

    ' Appearance Page
    .AlwaysOnTop = CBool(chkAlwaysOnTop.Value)
    .AutoHide = CBool(chkAutohide.Value)
    If optShowEntry.Value Then
      .TaskEntry = abtShow
    ElseIf optHideEntry.Value Then
      .TaskEntry = abtHide
    ElseIf optDependEntry.Value Then
      .TaskEntry = abtFloatDependent
    End If

    ' Position Page
    If optEdgeFloat.Value Then
      .Edge = abeFloat
    ElseIf optEdgeLeft.Value Then
      .Edge = abeLeft
    ElseIf optEdgeTop.Value Then
      .Edge = abeTop
    ElseIf optEdgeRight.Value Then
      .Edge = abeRight
    ElseIf optEdgeBottom.Value Then
      .Edge = abeBottom
    End If
  
    If chkAllowFloat.Value Then
      .Flags = .Flags Or abfAllowFloat
    Else
      .Flags = .Flags And (Not abfAllowFloat)
    End If
  
    If chkAllowLeft.Value Then
      .Flags = .Flags Or abfAllowLeft
    Else
      .Flags = .Flags And (Not abfAllowLeft)
    End If

    If chkAllowTop.Value Then
      .Flags = .Flags Or abfAllowTop
    Else
      .Flags = .Flags And (Not abfAllowTop)
    End If

    If chkAllowRight.Value Then
      .Flags = .Flags Or abfAllowRight
    Else
      .Flags = .Flags And (Not abfAllowRight)
    End If

    If chkAllowBottom.Value Then
      .Flags = .Flags Or abfAllowBottom
    Else
      .Flags = .Flags And (Not abfAllowBottom)
    End If

    ' Sizing Page
    .HorzSizeInc = updHorzSizeInc.Value
    .VertSizeInc = updVertSizeInc.Value

    ' Docking Page
    .MinHorzDockSize = updMinHorzDockSize.Value
    .MinVertDockSize = updMinVertDockSize.Value
    .HorzDockSize = updHorzDockSize.Value
    .VertDockSize = updVertDockSize.Value
    .MaxHorzDockSize = updMaxHorzDockSize.Value
    .MaxVertDockSize = updMaxVertDockSize.Value

    ' Floating Page
    .FloatLeft = updFloatLeft.Value
    .FloatTop = updFloatTop.Value
    .FloatRight = updFloatRight.Value
    .FloatBottom = updFloatBottom.Value
    
    .MinWidth = updMinWidth.Value
    .MinHeight = updMinHeight.Value
    .MaxWidth = updMaxWidth.Value
    .MaxHeight = updMaxHeight.Value

    ' Sliding Page
    .SlideEffect = CBool(chkSlideEffect.Value)
    .SlideTime = sldSlideTime.Value
    
    ' Registry Page
    If optCurrentUser.Value Then
      .RootKey = HKEY_CURRENT_USER
    ElseIf optLocalMachine.Value Then
      .RootKey = HKEY_LOCAL_MACHINE
    End If
    .KeyName = edtKeyName.Text
    
  End With

End Sub

Private Sub btnLoad_Click()
  
  Dim bSuccess As Boolean
  
  With frmChat.AppBar
  
    ' Set RootKey and KeyName properties
    If optCurrentUser.Value Then
      .RootKey = HKEY_CURRENT_USER
    ElseIf optLocalMachine.Value Then
      .RootKey = HKEY_LOCAL_MACHINE
    End If
    .KeyName = edtKeyName.Text
    
    ' Load settings from the registry
    bSuccess = .LoadSettings
    
    ' Show the operation result
    If bSuccess Then
      MsgBox "Settings successfully loaded.", vbInformation
    Else
      MsgBox "Failed to load settings.", vbExclamation
    End If
    
    ' Re-init dialog
    InitDialog
  
  End With
  
End Sub

Private Sub btnSave_Click()
  
  Dim bSuccess As Boolean
  
  With frmChat.AppBar
  
    ' Set RootKey and KeyName properties
    If optCurrentUser.Value Then
      .RootKey = HKEY_CURRENT_USER
    ElseIf optLocalMachine.Value Then
      .RootKey = HKEY_LOCAL_MACHINE
    End If
    .KeyName = edtKeyName.Text
    
    ' Save settings into the registry
    bSuccess = .SaveSettings
    
    ' Show the operation result
    If bSuccess Then
      MsgBox "Settings successfully saved.", vbInformation
    Else
      MsgBox "Failed to save settings.", vbExclamation
    End If
    
  End With
  
End Sub

Private Sub Form_Load()
  
  ' Initialize Dialog
  InitDialog
  
  ' Setup Tabbed Dialog
  For i = 0 To tabPage.Count - 1
    With tabPage(i)
      .Move tabMain.ClientLeft, _
            tabMain.ClientTop, _
            tabMain.ClientWidth, _
            tabMain.ClientHeight
    End With
  Next i
  
  tabPage(0).ZOrder 0
  
  For i = 1 To tabPage.Count - 1
    tabPage(i).Enabled = False
  Next i

End Sub

Private Sub btnApply_Click()
  
  ApplyChanges
  Unload frmAppBar

End Sub

Private Sub btnCancel_Click()
  
  Unload Me

End Sub

Private Sub lblEMail_Click()

End Sub

Private Sub tabMain_Click()
  
  For i = 0 To tabPage.Count - 1
    If i = tabMain.SelectedItem.Index - 1 Then
      tabPage(i).Enabled = True
    Else
      tabPage(i).Enabled = False
    End If
  Next i

  tabPage(tabMain.SelectedItem.Index - 1).ZOrder 0

End Sub

Private Sub edtHorzSizeInc_LostFocus()
  updHorzSizeInc.Value = CLng(edtHorzSizeInc.Text)
End Sub

Private Sub edtVertSizeInc_LostFocus()
  updVertSizeInc.Value = CLng(edtVertSizeInc.Text)
End Sub

Private Sub edtHorzDockSize_LostFocus()
  updHorzDockSize.Value = CLng(edtHorzDockSize.Text)
End Sub

Private Sub edtVertDockSize_LostFocus()
  updVertDockSize.Value = CLng(edtVertDockSize.Text)
End Sub

Private Sub edtFloatLeft_LostFocus()
  updFloatLeft.Value = CLng(edtFloatLeft.Text)
End Sub

Private Sub edtFloatTop_LostFocus()
  updFloatTop.Value = CLng(edtFloatTop.Text)
End Sub

Private Sub edtFloatRight_LostFocus()
  updFloatRight.Value = CLng(edtFloatRight.Text)
End Sub

Private Sub edtFloatBottom_LostFocus()
  updFloatBottom.Value = CLng(edtFloatBottom.Text)
End Sub

Private Sub edtMinWidth_LostFocus()
  updMinWidth.Value = CLng(edtMinWidth.Text)
End Sub

Private Sub edtMinHeight_LostFocus()
  updMinHeight.Value = CLng(edtMinHeight.Text)
End Sub

Private Sub edtMaxWidth_LostFocus()
  updMaxWidth.Value = CLng(edtMaxWidth.Text)
End Sub

Private Sub edtMaxHeight_LostFocus()
  updMaxHeight.Value = CLng(edtMaxHeight.Text)
End Sub

Private Sub edtMinVertDockSize_LostFocus()
  updMinVertDockSize.Value = CLng(edtMinVertDockSize.Text)
End Sub

Private Sub edtMaxVertDockSize_LostFocus()
  updMaxVertDockSize.Value = CLng(edtMaxVertDockSize.Text)
End Sub

Private Sub edtMinHorzDockSize_LostFocus()
  updMinHorzDockSize.Value = CLng(edtMinHorzDockSize.Text)
End Sub

Private Sub edtMaxHorzDockSize_LostFocus()
  updMaxHorzDockSize.Value = CLng(edtMaxHorzDockSize.Text)
End Sub

Private Sub chkSlideEffect_Click()
  sldSlideTime.Enabled = CBool(chkSlideEffect.Value)
  lblSlideTime.Enabled = CBool(chkSlideEffect.Value)
  lblFaster.Enabled = CBool(chkSlideEffect.Value)
  lblSlower.Enabled = CBool(chkSlideEffect.Value)
End Sub
