VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fBasic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 1 : Basic cTreeview Class Functions"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fBasic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDialog 
      Caption         =   "Control Treeview Using Code Only:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Index           =   3
      Left            =   6090
      TabIndex        =   69
      Top             =   12180
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Collapse  Node"
         Height          =   315
         Index           =   18
         Left            =   315
         TabIndex        =   86
         Top             =   4305
         Width           =   1350
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Expand Node"
         Height          =   315
         Index           =   17
         Left            =   315
         TabIndex        =   84
         Top             =   1260
         Width           =   1350
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Collapse All Child Nodes"
         Height          =   525
         Index           =   15
         Left            =   315
         TabIndex        =   87
         Top             =   4620
         Width           =   1350
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Expand All Nodes"
         Height          =   525
         Index           =   14
         Left            =   4200
         TabIndex        =   85
         Top             =   735
         Width           =   1350
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Expand All Child Nodes"
         Height          =   525
         Index           =   13
         Left            =   315
         TabIndex        =   83
         Top             =   735
         Width           =   1350
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Reset"
         Height          =   315
         Index           =   12
         Left            =   2520
         TabIndex        =   70
         Top             =   2730
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Top"
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   77
         Top             =   1050
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Page Up"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   78
         Top             =   1470
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Line Up"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   79
         Top             =   1890
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Line Down"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   80
         Top             =   3780
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Page Down"
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   81
         Top             =   4200
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Bottom"
         Height          =   315
         Index           =   5
         Left            =   2520
         TabIndex        =   82
         Top             =   4620
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Left"
         Height          =   315
         Index           =   6
         Left            =   315
         TabIndex        =   71
         Top             =   2730
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Page Left"
         Height          =   315
         Index           =   7
         Left            =   1470
         TabIndex        =   72
         Top             =   2415
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Line Left"
         Height          =   315
         Index           =   8
         Left            =   1470
         TabIndex        =   73
         Top             =   3045
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Line Right"
         Height          =   315
         Index           =   9
         Left            =   3570
         TabIndex        =   75
         Top             =   3045
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Page Right"
         Height          =   315
         Index           =   10
         Left            =   3570
         TabIndex        =   74
         Top             =   2415
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Right"
         Height          =   315
         Index           =   11
         Left            =   4620
         TabIndex        =   76
         Top             =   2730
         Width           =   930
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "Collapse All Nodes"
         Height          =   525
         Index           =   16
         Left            =   4200
         TabIndex        =   88
         Top             =   4620
         Width           =   1350
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Special Features:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Index           =   1
      Left            =   105
      TabIndex        =   147
      Top             =   12180
      Width           =   5895
      Begin VB.Frame fraBackground 
         Caption         =   "Background: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   56
         Top             =   1680
         Width           =   5265
         Begin VB.OptionButton optBackground 
            Appearance      =   0  'Flat
            Caption         =   "Custom Background color (Default color is vbWindowBackground )"
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   1
            Left            =   210
            TabIndex        =   170
            Top             =   650
            Width           =   3165
         End
         Begin VB.CommandButton cmdBackground 
            Caption         =   "..."
            Height          =   275
            Left            =   4725
            TabIndex        =   60
            Top             =   1260
            Width           =   435
         End
         Begin VB.OptionButton optBackground 
            Appearance      =   0  'Flat
            Caption         =   "Default (System Window Background color [ = vbWindowBackground] )"
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   0
            Left            =   210
            TabIndex        =   57
            Top             =   210
            Width           =   3270
         End
         Begin VB.PictureBox picBackground 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   3570
            Picture         =   "fBasic.frx":27A2
            ScaleHeight     =   960
            ScaleWidth      =   1590
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "Click to view original size"
            Top             =   210
            Width           =   1590
         End
         Begin VB.OptionButton optBackground 
            Appearance      =   0  'Flat
            Caption         =   "Custom Tiled Bitmap Image (Avoid image colors that clash with node colors)"
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   2
            Left            =   210
            TabIndex        =   58
            Top             =   1090
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.Label lblBackColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   3570
            TabIndex        =   171
            Top             =   210
            Visible         =   0   'False
            Width           =   1590
         End
      End
      Begin VB.CheckBox chkTooltips 
         Appearance      =   0  'Flat
         Caption         =   "Treeview Tooltips?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3780
         TabIndex        =   51
         Top             =   735
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame fraBorder 
         Caption         =   "Border Style: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   52
         Top             =   1050
         Width           =   5265
         Begin VB.OptionButton optBorderStyle 
            Appearance      =   0  'Flat
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   1365
            TabIndex        =   53
            Top             =   210
            Width           =   855
         End
         Begin VB.OptionButton optBorderStyle 
            Appearance      =   0  'Flat
            Caption         =   "Fixed Single"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   2520
            TabIndex        =   54
            Top             =   210
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optBorderStyle 
            Appearance      =   0  'Flat
            Caption         =   "Thin"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   3990
            TabIndex        =   55
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame fraOverlay 
         Caption         =   "Node Icon Overlay: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   62
         Top             =   4200
         Width           =   5265
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "Custom overlay icon 3"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   2625
            TabIndex        =   68
            Top             =   903
            Width           =   2235
         End
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "Custom overlay icon 2"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2625
            TabIndex        =   67
            Top             =   609
            Width           =   2235
         End
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "Custom overlay icon 1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2625
            TabIndex        =   66
            Top             =   315
            Width           =   2235
         End
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "Shortcut overlay icon"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   525
            TabIndex        =   65
            Top             =   903
            Width           =   2235
         End
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "Share overlay icon"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   525
            TabIndex        =   64
            Top             =   609
            Width           =   2235
         End
         Begin VB.OptionButton optOverlay 
            Appearance      =   0  'Flat
            Caption         =   "No overlay icon"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   525
            TabIndex        =   63
            Top             =   315
            Value           =   -1  'True
            Width           =   2235
         End
      End
      Begin VB.CheckBox chkEnabled 
         Appearance      =   0  'Flat
         Caption         =   "Treeview Eabled?"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2100
         TabIndex        =   50
         Top             =   735
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkLocked 
         Appearance      =   0  'Flat
         Caption         =   "Treeview Locked?"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   420
         TabIndex        =   49
         Top             =   735
         Width           =   1590
      End
      Begin VB.CheckBox chkCut 
         Appearance      =   0  'Flat
         Caption         =   "Selected Cut Icon State"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   61
         Top             =   3885
         Width           =   2115
      End
      Begin VB.Label lblFeature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Node"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   420
         TabIndex        =   151
         Top             =   3510
         Width           =   750
      End
      Begin VB.Label lblFeature 
         Alignment       =   2  'Center
         Caption         =   "Node"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Index           =   3
         Left            =   435
         TabIndex        =   150
         Top             =   3525
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         Index           =   2
         X1              =   210
         X2              =   5500
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Label lblFeature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Treeview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   420
         TabIndex        =   149
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label lblFeature 
         Alignment       =   2  'Center
         Caption         =   "Treeview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Index           =   0
         Left            =   435
         TabIndex        =   148
         Top             =   330
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   5
         Index           =   0
         X1              =   210
         X2              =   5500
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         Index           =   1
         X1              =   225
         X2              =   5525
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         Index           =   3
         X1              =   225
         X2              =   5525
         Y1              =   3690
         Y2              =   3690
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Node Details: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Index           =   0
      Left            =   105
      TabIndex        =   146
      Top             =   6825
      Width           =   5895
      Begin VB.CommandButton cmdVisTest 
         Caption         =   "Test"
         Height          =   300
         Left            =   5145
         TabIndex        =   169
         ToolTipText     =   "Try selecting a node, scrolling the node up/Down using the treeview control's scrollbars and click this button"
         Top             =   5460
         Width           =   540
      End
      Begin VB.OptionButton optDetail 
         Appearance      =   0  'Flat
         Caption         =   "Hover Item"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   6
         Top             =   315
         Width           =   1380
      End
      Begin VB.OptionButton optDetail 
         Appearance      =   0  'Flat
         Caption         =   "Selected Item"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Path Xpanded: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   13
         Left            =   2940
         TabIndex        =   168
         ToolTipText     =   "Is the Node's path fully expanded?"
         Top             =   3045
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   13
         Left            =   4305
         TabIndex        =   167
         Top             =   3045
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Child Count: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   12
         Left            =   105
         TabIndex        =   166
         Top             =   5145
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   12
         Left            =   1470
         TabIndex        =   165
         Top             =   5145
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vis ABS Pos: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   11
         Left            =   2940
         TabIndex        =   164
         ToolTipText     =   "Visible Absolute Position"
         Top             =   5460
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   11
         Left            =   4305
         TabIndex        =   163
         Top             =   5460
         Width           =   750
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "ABS Position: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   10
         Left            =   2940
         TabIndex        =   162
         ToolTipText     =   "Absolute Position"
         Top             =   5145
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   10
         Left            =   4305
         TabIndex        =   161
         Top             =   5145
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Position: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   9
         Left            =   105
         TabIndex        =   160
         Top             =   5460
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   9
         Left            =   1470
         TabIndex        =   159
         Top             =   5460
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "ABS Index: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   8
         Left            =   2940
         TabIndex        =   158
         ToolTipText     =   "Absolute Index"
         Top             =   4830
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   8
         Left            =   4305
         TabIndex        =   157
         Top             =   4830
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   7
         Left            =   1470
         TabIndex        =   48
         Top             =   4830
         Width           =   1380
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Text Unique: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   47
         Top             =   4830
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Full Key Path: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   0
         Left            =   1470
         TabIndex        =   34
         Top             =   3360
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Nesting level: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   35
         Top             =   3885
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "First Visible: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   37
         Top             =   4200
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Last Visible: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   4
         Left            =   105
         TabIndex        =   39
         Top             =   4515
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Root Node: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   5
         Left            =   2940
         TabIndex        =   41
         Top             =   3885
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Overlay Img: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   6
         Left            =   2940
         TabIndex        =   43
         ToolTipText     =   "Overlay Image ID"
         Top             =   4200
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Cut State: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   7
         Left            =   2940
         TabIndex        =   45
         ToolTipText     =   "Is node cut?"
         Top             =   4515
         Width           =   1275
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   36
         Top             =   3885
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   38
         Top             =   4200
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   1470
         TabIndex        =   40
         Top             =   4515
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   4305
         TabIndex        =   42
         Top             =   3885
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   4305
         TabIndex        =   44
         Top             =   4200
         Width           =   1380
      End
      Begin VB.Label lblTests 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   6
         Left            =   4305
         TabIndex        =   46
         Top             =   4515
         Width           =   1380
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         Caption         =   "Event Type: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   15
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Text: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Back Colour: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Bold: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Checked: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   15
         Top             =   2100
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Has Children: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   4
         Left            =   105
         TabIndex        =   17
         Top             =   2415
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Expanded: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   2730
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Fore Colour: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   6
         Left            =   105
         TabIndex        =   21
         Top             =   3045
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Full Path: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   7
         Left            =   105
         TabIndex        =   9
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Index: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   8
         Left            =   2940
         TabIndex        =   23
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Key: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   9
         Left            =   2940
         TabIndex        =   25
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Selected: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   12
         Left            =   2940
         TabIndex        =   27
         Top             =   2100
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Sorted: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   13
         Left            =   2940
         TabIndex        =   29
         Top             =   2415
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Visible: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   14
         Left            =   2940
         TabIndex        =   31
         Top             =   2730
         Width           =   1275
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   8
         Top             =   630
         Width           =   4215
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   12
         ToolTipText     =   "Back color"
         Top             =   1470
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   14
         ToolTipText     =   "Bold state condition"
         Top             =   1785
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   1470
         TabIndex        =   16
         ToolTipText     =   "Check state condition"
         Top             =   2100
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1470
         TabIndex        =   18
         ToolTipText     =   "Does node have children"
         Top             =   2415
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   1470
         TabIndex        =   20
         ToolTipText     =   "Is node expanded"
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   6
         Left            =   1470
         TabIndex        =   22
         ToolTipText     =   "Fore Color"
         Top             =   3045
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   7
         Left            =   1470
         TabIndex        =   10
         Top             =   945
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   8
         Left            =   4305
         TabIndex        =   24
         Top             =   1470
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   9
         Left            =   4305
         TabIndex        =   26
         Top             =   1785
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   12
         Left            =   4305
         TabIndex        =   28
         Top             =   2100
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   13
         Left            =   4305
         TabIndex        =   30
         Top             =   2415
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   14
         Left            =   4305
         TabIndex        =   32
         Top             =   2730
         Width           =   1380
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Move Nodes within own siblings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Index           =   4
      Left            =   11655
      TabIndex        =   152
      Top             =   14175
      Width           =   5895
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Last Node Position"
         Height          =   540
         Index           =   3
         Left            =   1890
         TabIndex        =   156
         Top             =   3675
         Width           =   2115
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Down One Node Position"
         Height          =   540
         Index           =   1
         Left            =   1890
         TabIndex        =   155
         Top             =   2835
         Width           =   2115
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Up one Node Position"
         Height          =   540
         Index           =   0
         Left            =   1890
         TabIndex        =   154
         Top             =   1995
         Width           =   2115
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "&First Node Position"
         Height          =   540
         Index           =   2
         Left            =   1890
         TabIndex        =   153
         Top             =   1155
         Width           =   2115
      End
   End
   Begin MSComctlLib.TabStrip tabSubs 
      Height          =   5160
      Left            =   6090
      TabIndex        =   89
      Top             =   6930
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9102
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.Tag             =   "0"
            Object.ToolTipText     =   "Add a new node"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.Tag             =   "0"
            Object.ToolTipText     =   "Edit Node Text and properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Copy a Node"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "Delete Node from Treeview"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            Object.Tag             =   "3"
            Object.ToolTipText     =   "Load, Save & clear the TreeView Control"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   3
      Left            =   15435
      ScaleHeight     =   5475
      ScaleWidth      =   5685
      TabIndex        =   133
      Top             =   12075
      Visible         =   0   'False
      Width           =   5685
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3675
         Width           =   4005
      End
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1470
         TabIndex        =   142
         ToolTipText     =   "NOTE: No check is done for valid characters."
         Top             =   3255
         Width           =   4005
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Go"
         Height          =   330
         Left            =   4410
         TabIndex        =   145
         Top             =   4200
         Width           =   1065
      End
      Begin VB.Frame fraFile 
         Caption         =   "Operation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Index           =   1
         Left            =   1470
         TabIndex        =   134
         Top             =   315
         Width           =   4005
         Begin VB.OptionButton optFileOp 
            Caption         =   "&Save To File"
            Height          =   330
            Index           =   2
            Left            =   1050
            TabIndex        =   137
            Top             =   945
            Value           =   -1  'True
            Width           =   2430
         End
         Begin VB.OptionButton optFileOp 
            Caption         =   "&Load from File"
            Height          =   330
            Index           =   1
            Left            =   1050
            TabIndex        =   136
            Top             =   630
            Width           =   2430
         End
         Begin VB.OptionButton optFileOp 
            Caption         =   "&Clear Treeview"
            Height          =   330
            Index           =   0
            Left            =   1050
            TabIndex        =   135
            Top             =   315
            Width           =   2430
         End
      End
      Begin VB.Frame fraFile 
         Caption         =   "File Format: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   0
         Left            =   1470
         TabIndex        =   138
         Top             =   1890
         Width           =   4005
         Begin VB.OptionButton optFileFmt 
            Caption         =   "&XML (Using ADO && XML 3.0)"
            Height          =   330
            Index           =   1
            Left            =   1050
            TabIndex        =   140
            Top             =   630
            Width           =   2430
         End
         Begin VB.OptionButton optFileFmt 
            Caption         =   "&Binary (smallest && Fastest)"
            Height          =   330
            Index           =   0
            Left            =   1050
            TabIndex        =   139
            Top             =   315
            Value           =   -1  'True
            Width           =   2430
         End
      End
      Begin VB.Label lblFile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "App Path:"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   143
         Top             =   3675
         Width           =   1275
      End
      Begin VB.Label lblFile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "FileName:"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   141
         Top             =   3255
         Width           =   1275
      End
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   2
      Left            =   15435
      ScaleHeight     =   5475
      ScaleWidth      =   5685
      TabIndex        =   127
      Top             =   6510
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CheckBox chkDelete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Include child nodes in delete operation:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1890
         TabIndex        =   131
         Top             =   2100
         Width           =   3165
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Go"
         Height          =   330
         Left            =   3990
         TabIndex        =   132
         Top             =   2625
         Width           =   1065
      End
      Begin VB.TextBox txtDelete 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1680
         Width           =   2850
      End
      Begin VB.Label lblInstruction 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "To select nodes, click on a node on the treeview control."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1575
         TabIndex        =   128
         Top             =   735
         Width           =   2850
      End
      Begin VB.Label lblDelete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete Node:"
         Height          =   315
         Left            =   840
         TabIndex        =   129
         Top             =   1680
         Width           =   1275
      End
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   1
      Left            =   15435
      ScaleHeight     =   5475
      ScaleWidth      =   5685
      TabIndex        =   119
      Top             =   945
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CheckBox chkCopy 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Include Child Nodes:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1995
         TabIndex        =   125
         Top             =   2625
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.TextBox txtCopy 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   1680
         Width           =   2850
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Go"
         Height          =   330
         Left            =   3990
         TabIndex        =   126
         Top             =   2625
         Width           =   1065
      End
      Begin VB.TextBox txtCopy 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   2100
         Width           =   2850
      End
      Begin VB.Label lblInstruction 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "To select nodes, click the text field below and then on a node on the treeview control."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   1575
         TabIndex        =   120
         Top             =   735
         Width           =   2850
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Destination node:"
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   123
         Top             =   2100
         Width           =   1275
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Node to copy:"
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   121
         Top             =   1680
         Width           =   1275
      End
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   0
      Left            =   9555
      ScaleHeight     =   5475
      ScaleWidth      =   5685
      TabIndex        =   90
      Top             =   945
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame fraSubs 
         Caption         =   "Sample Appearance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2940
         TabIndex        =   114
         Top             =   3465
         Width           =   2535
         Begin VB.PictureBox picSample 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   210
            ScaleHeight     =   540
            ScaleWidth      =   2115
            TabIndex        =   116
            Top             =   315
            Width           =   2115
            Begin VB.Label lblSample 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Sample Text"
               ForeColor       =   &H80000008&
               Height          =   220
               Left            =   510
               TabIndex        =   115
               Top             =   120
               Width           =   1380
            End
            Begin VB.Image imgSample 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   210
               Stretch         =   -1  'True
               Top             =   105
               Width           =   270
            End
         End
      End
      Begin VB.PictureBox picSubs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3675
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   103
         Top             =   1155
         Width           =   315
      End
      Begin VB.PictureBox picSubs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1155
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   100
         Top             =   1155
         Width           =   315
      End
      Begin VB.CommandButton cmdSubs 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   104
         Top             =   1155
         Width           =   435
      End
      Begin VB.CommandButton cmdSubs 
         Caption         =   "..."
         Height          =   330
         Index           =   0
         Left            =   1575
         TabIndex        =   101
         Top             =   1155
         Width           =   435
      End
      Begin VB.CheckBox chkSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Selected:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   2100
         TabIndex        =   106
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CheckBox chkSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Visible:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   4095
         TabIndex        =   107
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CheckBox chkSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Bold:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   105
         Top             =   1575
         Width           =   1275
      End
      Begin VB.ComboBox cboSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   2835
         Width           =   2430
      End
      Begin VB.ComboBox cboSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   2415
         Width           =   2430
      End
      Begin VB.ComboBox cboSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   1995
         Width           =   2430
      End
      Begin VB.ComboBox cboSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   3570
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   315
         Width           =   2010
      End
      Begin VB.CommandButton cmdSubs 
         Caption         =   "&Add"
         Height          =   330
         Index           =   2
         Left            =   3255
         TabIndex        =   117
         Top             =   4935
         Width           =   1065
      End
      Begin VB.TextBox txtSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1155
         TabIndex        =   92
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton cmdSubs 
         Caption         =   "&Cancel"
         Height          =   330
         Index           =   3
         Left            =   4410
         TabIndex        =   118
         Top             =   4935
         Width           =   1065
      End
      Begin VB.TextBox txtSubs 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1155
         TabIndex        =   94
         Top             =   735
         Width           =   1380
      End
      Begin VB.Label lblSubs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   3570
         TabIndex        =   98
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label lblSubs 
         Caption         =   "Relative:"
         Height          =   210
         Index           =   3
         Left            =   2625
         TabIndex        =   97
         Top             =   787
         Width           =   960
      End
      Begin VB.Image imgSubs 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   1260
         Stretch         =   -1  'True
         Top             =   2850
         Width           =   270
      End
      Begin VB.Image imgSubs 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   1260
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   270
      End
      Begin VB.Image imgSubs 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   1260
         Stretch         =   -1  'True
         Top             =   2010
         Width           =   270
      End
      Begin VB.Label lblSubs 
         Caption         =   "Xpand Image:"
         Height          =   210
         Index           =   9
         Left            =   105
         TabIndex        =   112
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "Sel Image:"
         Height          =   210
         Index           =   8
         Left            =   105
         TabIndex        =   110
         Top             =   2460
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "Norm Image:"
         Height          =   210
         Index           =   7
         Left            =   105
         TabIndex        =   108
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "BackColor:"
         Height          =   210
         Index           =   6
         Left            =   2625
         TabIndex        =   102
         Top             =   1215
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "ForeColor:"
         Height          =   210
         Index           =   5
         Left            =   105
         TabIndex        =   99
         Top             =   1215
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "Relationship:"
         Height          =   210
         Index           =   2
         Left            =   2625
         TabIndex        =   95
         Top             =   315
         Width           =   960
      End
      Begin VB.Label lblSubs 
         Caption         =   "Key:"
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   93
         Top             =   787
         Width           =   1065
      End
      Begin VB.Label lblSubs 
         Caption         =   "Text:"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   91
         Top             =   367
         Width           =   1065
      End
   End
   Begin MSComctlLib.TabStrip tabDialog 
      Height          =   6315
      Left            =   3465
      TabIndex        =   3
      Top             =   105
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11139
      TabWidthStyle   =   1
      Style           =   2
      TabFixedWidth   =   1939
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Object.ToolTipText     =   "Standard Node Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Features"
            Object.ToolTipText     =   "New features added to expand functionality"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Subs"
            Object.ToolTipText     =   "Various TreeView and Node functions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Scroll"
            Object.ToolTipText     =   "Scroll the Treeview's contents progmatically"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Move"
            Object.ToolTipText     =   "Shuffle a Node Up/Down own branch level"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkDialog 
      Appearance      =   0  'Flat
      Caption         =   "Allow Label Edit"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3570
      TabIndex        =   1
      Top             =   6510
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.CheckBox chkDialog 
      Appearance      =   0  'Flat
      Caption         =   "Hot Tracking Cursor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   5145
      TabIndex        =   2
      Top             =   6510
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   630
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483628
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":5228
            Key             =   "Closed1"
            Object.Tag             =   "Closed Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":57C2
            Key             =   "Open1"
            Object.Tag             =   "Open Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":5D5C
            Key             =   "Selected"
            Object.Tag             =   "Selected Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":62F6
            Key             =   "Group"
            Object.Tag             =   "Group Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":6890
            Key             =   "Closed2"
            Object.Tag             =   "Closed Network Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":6E2A
            Key             =   "Open2"
            Object.Tag             =   "Open Network Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":73C4
            Key             =   "Clock"
            Object.Tag             =   "Clock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":7816
            Key             =   "Barcode"
            Object.Tag             =   "Barcode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":7970
            Key             =   "Agent"
            Object.Tag             =   "Agent"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":9AAA
            Key             =   "Diary"
            Object.Tag             =   "Diary"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":F6CC
            Key             =   "Item"
            Object.Tag             =   "Card Item"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":F9E6
            Key             =   "ShareOverlay"
            Object.Tag             =   "Share Overlay"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":FF80
            Key             =   "ShortcutOverlay"
            Object.Tag             =   "Shortcut Overlay"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":100DA
            Key             =   "Custom1Overlay"
            Object.Tag             =   "Custom Overlay 1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":103F4
            Key             =   "Custom2Overlay"
            Object.Tag             =   "Custom Overlay 2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fBasic.frx":1054E
            Key             =   "Custom3Overlay"
            Object.Tag             =   "Custom Overlay 3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   11695
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "fBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fBasic [Tutorial 1]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         01/08/2002
' Version:      01.02.00
' Description:  Application showing basic features of vbTree DLL (cTreeView
'               class).
' Edit History: 01.00.00 23/03/2002 Initial Release
'               01.01.00 01/05/2002 Changed 'Move' tab to 'Scroll'; 'Move'
'                                   tab now shows how to move a node at
'                                   same branch level for own parent node;
'                                   Sample Treeview data changed to refect
'                                   fictional application; Fixed icon
'                                   comboboxes to allow for no icon when
'                                   adding/editing individual nodes
'               01.02.00 01/08/2002 Updated Properties Tab (Node Details)
'                                   to reflect new properties added in
'                                   version 02.02.00 of cTreeview Class.
'               01.02.00 01/08/2002 Updated tutorial to use cTabStrip class
'                                   to manage the two Tab controls.
'               01.03.00 15/08/2002 Updated project to demonstrate graphic
'                                   background images. This feature uniqely
'                                   maintains each individual node's display
'                                   properties - at this stage I haven't
'                                   found a successful way of stopping the
'                                   flickering when quickly scrolling with
'                                   the vertical scrollbar... Hmmm...
'
' Notes:        To view all GUI features, set form Height = 18000,
'               Width = 21500; to hide, set form Height = 7260,
'               Width = 9450
'
'===========================================================================

Option Explicit

#Const OLDDATA = 0      '## 1 = Old Test Data

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
#End If

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private mbIsBusy          As Boolean
Private moColor           As cColorDialog
Private mlNodeNdx         As Long
Private mlSelNdx          As Long
Private mlDestNdx         As Long
Private meNodeOp          As epicSub
Private meNodeFld         As etxtCopy
Private msFile            As String
Private msLastPath        As String     '@@ v01.03.00
Private moTabs(1 To 2)    As cTabStrip  '@@ v01.02.00

Private Enum eChkDialog
    [Allow Label Edit] = 0
    [Hot Tracking] = 1
End Enum

Private Enum eOptDetail
    [Selected Item] = 0
    [Hover Item] = 1
End Enum

Private Enum eCmdSubs
    [ForeColor] = 0
    [BackColor] = 1
    [Ok Command] = 2
    [Cancel Command] = 3
End Enum

Private Enum eChkSubs
    [IsBold] = 0
    [IsVisible] = 1
    [IsSelected] = 2
End Enum

Private Enum eImgCboSubs
    [Normal Image] = 0
    [Selected Image] = 2
    [Expanded Image] = 1
    [Relationships] = 3     '## **Combo Only**
End Enum

Private Enum eLblSubs
    [elblText] = 0
    [elblKey] = 1
    [elblRelationship] = 2
    [elblRelative] = 3
    [elblRelative Value] = 4
    [elblForeColor] = 5
    [elblBackColor] = 6
    [elblNormImage] = 7
    [elblSelImage] = 8
    [elblXpandImage] = 9
End Enum

Private Enum epicSub
    [epicEdit] = 0
    [epicCopy] = 1
    [epicDelete] = 2
    [epicLoad] = 3
End Enum

Private Enum eTxtSubs
    [eText] = 0
    [eKey] = 1
End Enum

Private Enum eInitSubs
    [Initalise] = 0
    [Add Node] = 1
    [Edit Node] = 2
    [Copy Node] = 3
    [Delete Node] = 4
    [Load Node] = 5
    [Clear Form] = 6
End Enum

Private Enum etxtCopy
    [Source Node] = 0
    [Dest Node] = 1
End Enum

Private Enum eoptFileOp
    efoclear = 0
    efoLoad = 1
    efoSave = 2
End Enum

Private Enum eCmdMove
    [eUp] = 0
    [eDown] = 1
    [eFirst] = 2
    [eLast] = 3
End Enum

Private Enum eCmdRemote
    [Top] = 0
    [Page Up] = 1
    [Line Up] = 2
    [Line Down] = 3
    [Page Down] = 4
    [Bottom] = 5

    [Left] = 6
    [Page Left] = 7
    [Line Left] = 8
    [Line Right] = 9
    [Page Right] = 10
    [Right] = 11

    [Reset] = 12

    [Expand Selected] = 17
    [Expand] = 13
    [Expand All] = 14

    [Collape Selected] = 18
    [Collape] = 15
    [Collapse All] = 16
End Enum

'===========================================================================
' Form Events
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    '## Make Return/Enter key act like the Tab key...
    '
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select

End Sub

Private Sub Form_Load()

#If NODLL = 0 Then
    Set moTree = New vbTree.cTreeView   '## Used to manage the TreeView
#Else
    Set moTree = New cTreeView
#End If

    Set moColor = New cColorDialog      '## Used to manage the color selection dialog

    pInitSubs [Initalise]               '## Initalises Add/Edit Subroutines

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog
        '
        '## Initialise the overlay icons
        '
        .SetOverlayImage ilDialog.hImageList, 11, [Share Overlay]
        .SetOverlayImage ilDialog.hImageList, 12, [Shortcut Overlay]
        .SetOverlayImage ilDialog.hImageList, 13, [Custom 1]
        .SetOverlayImage ilDialog.hImageList, 14, [Custom 2]
        .SetOverlayImage ilDialog.hImageList, 15, [Custom 3]
        '
        '## Disable the treeview's refresh for faster loading
        '
        .Redraw False
        '
        '## Set TreeView features
        '
        With .Ctrl
            .Style = tvwTreelinesPlusMinusPictureText
            .LineStyle = tvwRootLines
            .Indentation = 10
            .ImageList = ilDialog
            .FullRowSelect = False
            .HideSelection = False
            .HotTracking = True
            .LabelEdit = tvwAutomatic
            '
            '## Build TreeView data
            '
            pInitData
        End With
        '
        '## Initialise the selected node's properties
        '
        mlNodeNdx = 1
        tvwDialog.Nodes(mlNodeNdx).Selected = True
        pShowNodeDetails .Ctrl.SelectedItem
        '
        '@@ v01.03.00 {
        '
        '## Setup Backround details
        '
        .BackMode = bmGraphic
        .BackColor = vbInfoBackground
        lblBackColor.BackColor = vbInfoBackground
        Set .BackPicture = picBackground.Picture
        msLastPath = App.Path
        '}
        '
        '## Enable the treeview to display the nodes
        '
        .Redraw True
        '.Refresh
    End With
'@@ v01.02.00 {
''    '
''    '## Set the initial Tab <- causes a redraw
''    '
''    tabDialog.SelectedItem = tabDialog.Tabs(1)
    '
    '## Hook the TabStrip controls and associate controls with each Tab.
    '
    Set moTabs(1) = New cTabStrip
    With moTabs(1)
        .HookCtrl tabDialog
        .AutoFit = True
        .Attach 1, fraDialog(0)
        .Attach 2, fraDialog(1)
        .Attach 3, tabSubs
        .Attach 4, fraDialog(3)
        .Attach 5, fraDialog(4)
        Set .TabStipContainer = Me
    End With

    Set moTabs(2) = New cTabStrip
    With moTabs(2)
        .HookCtrl tabSubs
        .AutoFit = True
        .Attach 1, picSub([epicEdit])
        .Attach 2, picSub([epicEdit])
        .Attach 3, picSub([epicCopy])
        .Attach 4, picSub([epicDelete])
        .Attach 5, picSub([epicLoad])
        Set .TabStipContainer = Me
    End With
'}
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
    Set moColor = Nothing
End Sub

'===========================================================================
' Form Control Events
'
Private Sub chkDialog_Click(Index As Integer)

    Dim bState As Boolean

    bState = CBool(chkDialog(Index).Value = 1)  '## Get the clicked checkbox state
    With moTree
        Select Case Index                       '## Which checkbox was clicked?
            Case [Allow Label Edit]
                .Ctrl.LabelEdit = CByte(Abs(bState = False))
            Case [Hot Tracking]
                .Ctrl.HotTracking = bState
                '
                '## Toggle the Event type based on whether or not HotTracking
                '   is enabled/disabled
                '
                pToggleHoverOpt
        End Select
    End With
End Sub

Private Sub tabDialog_Click()   '@@ v01.02.00

    Select Case tabDialog.SelectedItem.Index
        Case 3: meNodeOp = 0
    End Select

''    Dim lPtr  As Long
''    Dim lLoop As Long
''
''    With tabDialog
''        For lLoop = 1 To .Tabs.Count
''            lPtr = lLoop - 1
''            Select Case lPtr
''                Case 2  '## Tab Subs doesn't use a frame control - special handling required
''                    '
''                    '## Was the tab selected?
''                    '
''                    If (lLoop = .SelectedItem.Index) Then
''                        '
''                        '## Yes.
''                        '
''                        tabSubs.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
''                        tabSubs.ZOrder                              '## Bring to front
''                        picSub([epicEdit]).Move tabSubs.ClientLeft, tabSubs.ClientTop, tabSubs.ClientWidth, tabSubs.ClientHeight
''                        picSub([epicEdit]).ZOrder                   '## Bring to front
''                        tabSubs_Click                               '## Prepare the 'Sub' Tab
''                    End If
''                    tabSubs.Visible = (.SelectedItem.Index = 3)     '## Show/Hide
''                Case Else
''                    '
''                    '## Position and show/hide the frame controls
''                    '
''                    fraDialog(lPtr).Visible = (lLoop = .SelectedItem.Index)
''                    fraDialog(lPtr).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
''                    fraDialog(lPtr).ZOrder                          '## Bring to front
''                    meNodeOp = 0
''            End Select
''        Next
''    End With

End Sub

'===========================================================================
' TAB 1 : Properties
'
Private Sub cmdVisTest_Click()
    lblTests(11).Caption = moTree.NodeAbsolutePosition(tvwDialog.SelectedItem, True)
End Sub

Private Sub moTree_AfterLabelEdit(Cancel As Integer, NewString As String)
    With moTree.Ctrl
        .SelectedItem.Text = NewString  '## NewString not commited until after AfterLabelEdit
                                        '   event - so manually commit the change.
        pShowNodeDetails .SelectedItem  '## Update the displayed node data
    End With
    pLoadFields
End Sub

Private Sub moTree_Hover(Node As MSComctlLib.Node)
    If optDetail([Hover Item]).Value Then
        '
        '## Only update the properties if event-type is hover
        '
        pShowNodeDetails Node           '## Update the displayed node data
    End If
End Sub

Private Sub moTree_SelChange()
    Debug.Print "Selection Change"
    pShowNodeDetails tvwDialog.SelectedItem
    pLoadFields                                 '## Update Add/Edit fields under Tab 'sub'
End Sub

Private Sub moTree_Selected(Node As MSComctlLib.Node)

    mlNodeNdx = Node.Index

    If meNodeOp > 0 Then
        '
        '## Node Copy, Delete operations (TAB 3: SUBS)
        '
        mbIsBusy = True
        Select Case meNodeOp
            Case [epicCopy]
                Select Case meNodeFld
                    Case [Source Node]
                        mlSelNdx = mlNodeNdx
                        txtCopy([Source Node]).Text = Node.Text
                    Case [Dest Node]
                        mlDestNdx = mlNodeNdx
                        txtCopy([Dest Node]).Text = Node.Text
                End Select
            Case [epicDelete]
                txtDelete.Text = Node.Text
        End Select
        mbIsBusy = False
    Else
        '
        '## Update Properties and Features Tab (1 & 2) fields
        '
        If optDetail([Selected Item]).Value Then
            '
            '## Only update the properties if event-type is selected
            '
            pShowNodeDetails Node           '## Update the displayed node data
            pLoadFields                     '## Update Add/Edit fields under Tab 'sub'
        End If
    End If

End Sub

Private Sub optDetail_Click(Index As Integer)
    Select Case Index
        Case [Selected Item]
            With moTree.Ctrl
                pShowNodeDetails .SelectedItem  '## Update the displayed node data
            End With
    End Select
End Sub

Private Sub pToggleHoverOpt()
    Select Case True
        Case (chkDialog([Hot Tracking]).Value = 0)      '## HotTracking is disabled
            optDetail([Selected Item]).Value = True     '## Set Event-type to selected
            optDetail([Hover Item]).Enabled = False     '## Disable hover event-type
        Case Else                                       '## HotTracking is enabled
            optDetail([Hover Item]).Enabled = True      '## Ensable hover event-type
    End Select
End Sub

Private Sub pShowNodeDetails(Node As MSComctlLib.Node)
    '
    '## Show all node details to user
    '
    Dim sText As String

    If (Node Is Nothing) Then Exit Sub
    If moTree.Locked Then Exit Sub

    mbIsBusy = True     '## Stop cascading events whilst changing a control's value
    LockWindowUpdate Me.hWnd
    With moTree
        optOverlay(.OverlayIcon(Node)).Value = True     '## Show a node's Overlay option
        chkCut.Value = Abs(CInt(.CutIcon(Node)))        '## Show the node's 'cut' state
    End With

    'fraDialog(0).Visible = False
    With Node
        lblDetails(0).Caption = .Text
        lblDetails(1).Caption = "&&H" + Hex$(.BackColor)
        lblDetails(2).Caption = IIf(.Bold, "Yes", "No")
        lblDetails(3).Caption = IIf(.Checked, "Yes", "No")
        If .Children Then
            sText = "Yes, " + CStr(.Children)
        Else
            sText = "No"
        End If
        lblDetails(4).Caption = sText
        lblDetails(5).Caption = IIf(.Expanded, "Yes", "No")
        lblDetails(6).Caption = "&&H" + Hex$(.ForeColor)
        lblDetails(7).Caption = "  \" + .FullPath
        lblDetails(8).Caption = CStr(.Index)
        lblDetails(9).Caption = .Key
        If moTree.IsRootNode(Node) Then
            sText = "Node is Root"
        Else
            sText = .Parent.Text
        End If
        lblDetails(12).Caption = IIf(.Selected, "Yes", "No")
        lblDetails(13).Caption = IIf(.Sorted, "Yes", "No")
        lblDetails(14).Caption = IIf(.Visible, "Yes", "No")

        lblTests(0).Caption = "  \" + moTree.FullKeyPath(Node)
        lblTests(1).Caption = moTree.NodeNestingLevel(Node)
        lblTests(2).Caption = moTree.FirstVisibleNode
        lblTests(3).Caption = moTree.LastVisibleNode
        lblTests(4).Caption = .Root.Text
        lblTests(5).Caption = moTree.OverlayIcon(Node)
        lblTests(6).Caption = moTree.CutIcon(Node)
        '@@ { cTreeView v02.02.00 New Properties
        lblTests(7).Caption = IIf(moTree.IsNodeTextUnique(Node, .Text, vbTextCompare), "Yes", "No")
        lblTests(8).Caption = moTree.NodeAbsoluteIndex(Node)
        lblTests(9).Caption = moTree.NodePosition(Node)
        lblTests(10).Caption = moTree.NodeAbsolutePosition(Node)
        lblTests(11).Caption = moTree.NodeAbsolutePosition(Node, True)
        lblTests(12).Caption = moTree.NodeCountChildren(Node, False, False)
        lblTests(13).Caption = IIf(moTree.IsPathExpanded(Node), "Yes", "No")
        '@@ }
    End With
    'fraDialog(0).Visible = True
    mbIsBusy = False    '## Enable Checkbox & Option events
    LockWindowUpdate 0

End Sub

'===========================================================================
' TAB 2 : Features
'
Private Sub chkCut_Click()
    If mbIsBusy Then Exit Sub
    '
    '## Set the node's Cut state
    '
    moTree.CutIcon = (chkCut.Value = vbChecked)
    pShowNodeDetails moTree.Ctrl.SelectedItem   '## Update the displayed node data
End Sub

Private Sub chkEnabled_Click()
    moTree.Enabled = (chkEnabled.Value = vbChecked)
End Sub

Private Sub chkLocked_Click()
    moTree.Locked = (chkLocked.Value = vbChecked)
End Sub

Private Sub chkTooltips_Click()
    moTree.ToolTips = (chkTooltips.Value = vbChecked)
End Sub

Private Sub optBorderStyle_Click(Index As Integer)
    moTree.Appearance = CByte(Index)
End Sub

Private Sub optOverlay_Click(Index As Integer)
    If mbIsBusy Then Exit Sub
    '
    '## Set the node's overlay icon
    '
    moTree.OverlayIcon(tvwDialog.SelectedItem) = CLng(Index)
    pShowNodeDetails moTree.Ctrl.SelectedItem   '## Update the displayed node data
End Sub

'---------------------------------------------------------------------------
' Background
'
'@@ v01.03.00
Private Sub cmdBackground_Click()
 
    Dim lResult   As Long
    Dim oDlg      As CFileDialog
    Dim sFileName As String

    Select Case True
        Case optBackground(bmColor).Value = True
            '
            '## Display Color Selector Dialog
            '
            With moColor
                .ColorFlags = CC_ANYCOLOR Or CC_RGBINIT
                .InitColor = moTree.BackColor
                lResult = .Execute(Me)
                If lResult Then
                    '
                    '## Update the displayed color with the user-selected
                    '
                    lblBackColor.BackColor = lResult
                    moTree.BackColor = lResult
                End If
            End With

        Case optBackground(bmGraphic).Value = True
            Set oDlg = New CFileDialog
            With oDlg
                .DialogMode = cdlgOpen
                .Path = msLastPath
                .Flags = cdlgOFNFileMustExist Or cdlgOFNHideReadOnly Or cdlgOFNExplorer
                .hwndOwner = Me.hWnd
                .Mask = "Graphics Files (*.Bmp, *.jpeg, *.jpg, *.gif)|*.Bmp;*.jpeg;*.jpg;*.gif"
                .Title = "Select Image"
                sFileName = .FileName
                msLastPath = .Path
            End With
            Set oDlg = Nothing
            If Len(Trim$(sFileName)) Then
                picBackground.Picture = LoadPicture(sFileName)
                Set moTree.BackPicture = picBackground.Picture
            End If

    End Select

End Sub

'@@ v01.03.00
Private Sub optBackground_Click(Index As Integer)
    cmdBackground.Visible = Not (Index = bmDefault)
    lblBackColor.Visible = (Index = bmColor)
    picBackground.Visible = (Index = bmGraphic)
    moTree.BackMode = CByte(Index)
End Sub

'@@ v01.03.00
Private Sub picBackground_Click()

    Dim oFrm As fViewBMP

    If optBackground(bmGraphic).Value = True Then
        Set oFrm = New fViewBMP
        Set oFrm.PicBox = picBackground
        oFrm.Show vbModal
        Set oFrm = Nothing
    End If

End Sub

'===========================================================================
' TAB 3 : Subs
'
Private Sub tabSubs_Click()     '@@ v01.02.00

    With tabSubs.SelectedItem
        meNodeOp = CByte(.Tag)
        pInitSubs CByte(.Index)
        If meNodeOp > epicEdit Then
            moTree_Selected tvwDialog.SelectedItem  '## Auto-load Copy, delete field when operation selected
        End If
    End With

''    Dim lLoop   As Long
''    Dim lPtr    As Long
''    Dim lPicPtr As Long
''
''    With tabSubs
''        For lLoop = 1 To .Tabs.Count                    '## Cycle through all tabs
''            lPicPtr = CStr(.Tabs(lLoop).Tag)            '## Extract control container ID
''            If (lLoop = .SelectedItem.Index) Then       '## Is it the selected tab?
''                pInitSubs lLoop                         '## Initialise the container
''                lPtr = lPicPtr                          '## Remember selected tab
''                meNodeOp = CByte(lPtr)                  '## Remember which node operation has focus
''                If lPtr > 0 Then
''                    moTree_Selected tvwDialog.SelectedItem  '## Auto-load Copy, delete field
''                                                            '   when operation selected
''                End If
''            End If
''            picSub(lPicPtr).Visible = (lPtr = lPicPtr)  '## Display only for selected tab
''            picSub(lPicPtr).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
''            picSub(lPicPtr).ZOrder                      '## Bring to front
''        Next
''    End With

End Sub

Private Sub pInitSubs(Mode As eInitSubs)

    Dim lLoop As Long
    Dim jLoop As Long

    Select Case Mode
        Case [Initalise]        '## Initalise all Add/Edit controls
            mbIsBusy = True     '## Stop cascading events whilst changing a control's value
            With ilDialog
                For jLoop = 0 To 3
                    With cboSubs(jLoop)
                        .Visible = False
                        .Clear
                        .Visible = True
                    End With
                Next
                For jLoop = 0 To 2                                  '@@ v01.01.00
                    cboSubs(jLoop).AddItem "[None]"
                Next
                For lLoop = 1 To .ListImages.Count
                    For jLoop = 0 To 2
                        cboSubs(jLoop).AddItem .ListImages(lLoop).Tag
                    Next
                Next
                With cboSubs([Relationships])
                    .AddItem "Child of Relative": .ItemData(.NewIndex) = 4
                    .AddItem "First":             .ItemData(.NewIndex) = 0
                    .AddItem "Before Relative":   .ItemData(.NewIndex) = 3
                    .AddItem "After Relative":    .ItemData(.NewIndex) = 2
                    .AddItem "Last":              .ItemData(.NewIndex) = 1
                    .ListIndex = 0
                End With
            End With
            mbIsBusy = False            '## Enable Checkbox & Option events
            pInitSubs [Clear Form]      '## Now clear & set defaults

        Case [Clear Form]               '## Clear form
            mbIsBusy = True     '## Stop cascading events whilst changing a control's value
            txtSubs([eText]).Text = ""
            txtSubs([eKey]).Text = ""
            picSubs([ForeColor]).BackColor = vbButtonText
            picSubs([BackColor]).BackColor = vbWindowBackground
            chkSubs([IsBold]).Value = vbUnchecked
            chkSubs([IsVisible]).Value = vbUnchecked
            chkSubs([IsSelected]).Value = vbUnchecked
            For jLoop = 0 To 2
                cboSubs(jLoop).ListIndex = 0
                imgSubs(jLoop) = Me.Picture                             '@@ v01.01.00
            Next
            mbIsBusy = False            '## Enable Checkbox & Option events
            pUpdateSampleDisplay        '## Update the sample node picture

        Case [Add Node]
            lblSubs([elblRelationship]).Visible = True
            lblSubs([elblRelative]).Visible = True
            lblSubs([elblRelative Value]).Visible = True
            cboSubs([Relationships]).Visible = True
            With txtSubs([eKey])
                .Locked = False
                .BackColor = vbWindowBackground
            End With
            cmdSubs([Ok Command]).Caption = "&Add"
            cmdSubs([Cancel Command]).Caption = "&Clear"
            pInitSubs [Clear Form]
            pLoadFields

        Case [Edit Node]
            lblSubs([elblRelationship]).Visible = False
            lblSubs([elblRelative]).Visible = False
            lblSubs([elblRelative Value]).Visible = False
            cboSubs([Relationships]).Visible = False
            With txtSubs([eKey])
                .Locked = True
                .BackColor = vbInfoBackground
            End With
            cmdSubs([Ok Command]).Caption = "&Update"
            cmdSubs([Cancel Command]).Caption = "&Cancel"
            pInitSubs [Clear Form]
            pLoadFields

        Case [Load Node]
            pInitFileOp

    End Select
End Sub

'---------------------------------------------------------------------------
' TAB 3 : Subs : Add/Edit Node
'---------------------------------------------------------------------------
'
Private Sub cboSubs_Click(Index As Integer)
    If mbIsBusy Then Exit Sub
    Select Case Index
        Case [Relationships]
            '## Do nothing ...
        Case Else
            '
            '## Icon changed, update the displayed icon
            '
            If cboSubs(Index).ListIndex Then
                imgSubs(Index) = ilDialog.ListImages(cboSubs(Index).ListIndex).Picture
            Else
                imgSubs(Index) = Me.Picture
            End If
            pUpdateSampleDisplay            '## Update the sample node picture
    End Select
End Sub

Private Sub chkSubs_Click(Index As Integer)
    pUpdateSampleDisplay                    '## Update the sample node picture
End Sub

Private Sub cmdSubs_Click(Index As Integer)

    Dim lResult As Long

    Select Case Index
        Case [ForeColor], [BackColor]
            '
            '## Display Color Selector Dialog
            '
            With moColor
                .ColorFlags = CC_ANYCOLOR Or CC_RGBINIT
                .InitColor = picSubs(Index).BackColor
                lResult = .Execute(picSubs(Index))
                If lResult Then
                    '
                    '## Update the displayed color with the user-selected
                    '
                    picSubs(Index).BackColor = lResult
                End If
            End With
            pUpdateSampleDisplay                '## Update the sample node picture

        Case [Ok Command]
            Select Case tabSubs.SelectedItem.Index
                Case 1: pAddNode                '## Add new user-designed node
                Case 2: pUpdateNode             '## Update the node with new properties
            End Select

        Case [Cancel Command]
            Select Case tabSubs.SelectedItem.Index
                Case 1: pInitSubs [Clear Form]  '## Cancel and reset fields
                Case 2: pLoadFields             '## Update the node with new properties
            End Select

    End Select

End Sub

Private Sub txtSubs_Change(Index As Integer)
    pUpdateSampleDisplay            '## Update the sample node picture
End Sub

Private Sub txtSubs_GotFocus(Index As Integer)
    pHiLite txtSubs(Index)
End Sub

Private Sub pAddNode()

    '
    '## Check compulsory Fields
    '
    Select Case True
        Case Len(Trim$(txtSubs([eText]).Text)) = 0
            MsgBox "Node must have a name", vbCritical, "ADD NODE"
            Exit Sub
        Case Len(Trim$(txtSubs([eKey]).Text)) = 0
            MsgBox "Node must have a Key", vbCritical, "ADD NODE"
            Exit Sub
        Case moTree.NodeExist(txtSubs([eKey]).Text) = True
            MsgBox "Node Key '" + txtSubs([eKey]).Text + "' already exists!", vbCritical, "ADD NODE"
            Exit Sub
    End Select

    '
    '## Add Node
    '
    With moTree
        .NodeAdd .Ctrl.Nodes(mlNodeNdx), _
                 cboSubs([Relationships]).ItemData(cboSubs([Relationships]).ListIndex), _
                 txtSubs([eKey]).Text, _
                 txtSubs([eText]).Text, _
                 cboSubs([Normal Image]).ListIndex, _
                 cboSubs([Selected Image]).ListIndex, , _
                 CBool(chkSubs([IsBold]).Value = vbChecked), , _
                 CBool(chkSubs([IsVisible]).Value = vbChecked), , _
                 CBool(chkSubs([IsSelected]).Value = vbChecked), , _
                 picSubs([ForeColor]).BackColor, _
                 picSubs([BackColor]).BackColor, _
                 cboSubs([Expanded Image]).ListIndex                        '@@ v01.01.00
    End With

End Sub

Private Sub pLoadFields()

    If moTree.Locked Then Exit Sub
    '
    '## Take data from selected node to Add/Edit Fields
    '
    With tvwDialog.SelectedItem
        Select Case tabSubs.SelectedItem.Index
            Case [Add Node]
                lblSubs([elblRelative Value]).Caption = .Text

            Case [Edit Node]
                txtSubs([eText]).Text = .Text
                txtSubs([eKey]).Text = .Key
                picSubs([ForeColor]).BackColor = .ForeColor
                picSubs([BackColor]).BackColor = .BackColor
                chkSubs([IsBold]).Value = Abs(CInt(.Bold))
                chkSubs([IsVisible]).Value = Abs(CInt(.Visible))
                chkSubs([IsSelected]).Value = Abs(CInt(.Selected))
                If .Image Then                                              '@@ v01.01.00
                    imgSubs([Normal Image]).Picture = ilDialog.ListImages(.Image).Picture
                Else                                                        '@@ v01.01.00
                    imgSubs([Normal Image]).Picture = Me.Picture            '@@
                End If                                                      '@@
                cboSubs([Normal Image]).ListIndex = .Image                  '@@
                If .SelectedImage Then                                      '@@ v01.01.00
                    imgSubs([Selected Image]).Picture = ilDialog.ListImages(.SelectedImage).Picture
                Else                                                        '@@
                    imgSubs([Normal Image]).Picture = Me.Picture            '@@ v01.01.00
                End If                                                      '@@
                cboSubs([Selected Image]).ListIndex = .SelectedImage        '@@
                If .ExpandedImage Then                                      '@@
                    imgSubs([Expanded Image]).Picture = ilDialog.ListImages(.ExpandedImage).Picture
                Else                                                        '@@ v01.01.00
                    imgSubs([Expanded Image]).Picture = Me.Picture          '@@
                End If                                                      '@@
                cboSubs([Expanded Image]).ListIndex = .ExpandedImage        '@@

        End Select
    End With
End Sub

Private Sub pUpdateNode()
    '
    '## Update selected node with edit field data
    '
    With moTree.Ctrl.Nodes(mlNodeNdx)
        .Text = txtSubs([eText]).Text
        .Image = cboSubs([Normal Image]).ListIndex                          '@@ v01.01.00
        .SelectedImage = cboSubs([Selected Image]).ListIndex                '@@
        .Bold = CBool(chkSubs([IsBold]).Value = vbChecked)
        If CBool(chkSubs([IsVisible]).Value = vbChecked) Then .EnsureVisible
        .Selected = CBool(chkSubs([IsSelected]).Value = vbChecked)
        .ForeColor = picSubs([ForeColor]).BackColor
        .BackColor = picSubs([BackColor]).BackColor
        .ExpandedImage = cboSubs([Expanded Image]).ListIndex                '@@ v01.01.00
    End With

End Sub

Private Sub pUpdateSampleDisplay()
    '
    '## Update sample node display
    '
    If chkSubs([IsSelected]).Value Then
        Set imgSample.Picture = imgSubs([Selected Image]).Picture
    Else
        Set imgSample.Picture = imgSubs([Normal Image]).Picture
    End If
    With lblSample
        If Len(Trim$(txtSubs([eText]).Text)) Then
            .Caption = " " + Replace(txtSubs([eText]).Text, "&", "&&")
        Else
            .Caption = " Sample Text"
        End If
        .ForeColor = picSubs([ForeColor]).BackColor
        .BackColor = picSubs([BackColor]).BackColor
        .FontBold = CBool(chkSubs([IsBold]).Value = vbChecked)
    End With
End Sub

'---------------------------------------------------------------------------
' TAB 3 : TAB Subs : Copy Node
'---------------------------------------------------------------------------
'
Private Sub cmdCopy_Click()
    pCopyNode
    pLoadFields
End Sub

Private Sub moTree_CopyNode(DestNode As MSComctlLib.INode, SrcNode As MSComctlLib.INode, Cancel As Boolean)

    '
    '## Raised by moTree.NodeCopy to do the physical operation. Cannot be done from
    '   within the class due to too many external factors. In this case the key
    '   is the same one used in the database with a prefix - therefore needs to
    '   be generated external to the cTreeView class.
    '
    Dim lID     As Long
    Dim sOldKey As String
    Dim sNewKey As String

    With SrcNode
        sOldKey = Replace(.Key, "[Copy]", "")
        '
        '## Find a unique key
        '
        Do
            sNewKey = sOldKey + CStr(lID) + "[Copy]"
            lID = lID + 1
            If Not moTree.NodeExist(sNewKey) Then
                Exit Do
            End If
        Loop
        '
        '## Add new node
        '
        moTree.NodeAdd DestNode, _
                       tvwChild, _
                       sNewKey, _
                       .Text, _
                       .Image, _
                       .SelectedImage, _
                       .Tag, _
                       .Bold, _
                       .Checked, , _
                       .Expanded, , _
                       .Sorted, _
                       .ForeColor, _
                       .BackColor, _
                       .ExpandedImage
    End With

End Sub

Private Sub txtCopy_GotFocus(Index As Integer)
    pHiLite txtCopy(Index)
    meNodeFld = CByte(Index)
End Sub

Private Sub pCopyNode()

    Dim oSelNode  As MSComctlLib.Node
    Dim oDestNode As MSComctlLib.Node

    If Len(txtCopy([Source Node]).Text) Then
        If Len(txtCopy([Dest Node]).Text) Then
            Set oSelNode = tvwDialog.Nodes(mlSelNdx)
            Set oDestNode = tvwDialog.Nodes(mlDestNdx)
            If chkCopy.Value = 1 Then
                '## Check for possible infinite looping and disable if true
                chkCopy.Value = Abs(CLng(moTree.IsParentNode(oDestNode, oSelNode) = False))
            End If
            tvwDialog.Visible = False                   '## Speed up drawing and stop flickering
            '
            '## Did we copy all required nodes?
            '
            If Not moTree.NodeCopy(oDestNode, oSelNode, CBool(chkCopy.Value = 1)) Then
                '
                '## No. Problem copying the node. Most likely an ADO error.
                '
                MsgBox "Unable to copy the selected node(s).", _
                       vbApplicationModal + vbExclamation + vbOKOnly, _
                       "Copy Node"
                txtCopy([Source Node]).SetFocus
                tvwDialog.SetFocus
            End If
        Else
            '
            '## No destination node selected
            '
            MsgBox "No Destination node selected.", _
                   vbApplicationModal + vbExclamation + vbOKOnly, _
                   "Copy Node"
            txtCopy([Dest Node]).SetFocus
        End If
    Else
        '
        '## No source node selected
        '
        MsgBox "No Source node selected.", _
               vbApplicationModal + vbExclamation + vbOKOnly, _
               "Copy Node"
        txtCopy([Source Node]).SetFocus
    End If
    tvwDialog.Visible = True

End Sub

'---------------------------------------------------------------------------
' TAB 3 : TAB Subs : Delete Node
'---------------------------------------------------------------------------
'
Private Sub cmdDelete_Click()
    pDeleteNode
End Sub

Private Sub txtDelete_GotFocus()
    pHiLite txtDelete
End Sub

Private Sub pDeleteNode()

    Dim oNode  As MSComctlLib.Node
    'Dim lCount As Long
    'Dim lLoop  As Long

    If Len(txtDelete.Text) Then
        '
        '## Just in case of manual user input
        '
        Set oNode = tvwDialog.Nodes(mlNodeNdx)
        If moTree.NodeExist(oNode.Key) Then
            If Not moTree.NodeDelete(oNode, True, CBool(chkDelete.Value = vbChecked)) Then
                MsgBox "Unable to delete the selected node.", _
                       vbApplicationModal + vbExclamation + vbOKOnly, _
                       "Delete Node"
                txtDelete.SetFocus
            End If
        Else
            MsgBox "Node does not exist.", _
                   vbApplicationModal + vbExclamation + vbOKOnly, _
                   "Delete Node"
            txtDelete.SetFocus
        End If
    Else
        '
        '## No source node selected
        '
        MsgBox "No Source node selected.", _
               vbApplicationModal + vbExclamation + vbOKOnly, _
               "Delete Node"
        txtDelete.SetFocus
    End If

End Sub

'---------------------------------------------------------------------------
' TAB 3 : Subs : Load/Save
'---------------------------------------------------------------------------
'
Private Sub optFileFmt_Click(Index As Integer)
'
End Sub

Private Sub optFileOp_Click(Index As Integer)

    Dim bState As Boolean

    bState = (optFileOp(efoclear).Value = False)
    fraFile(0).Enabled = bState
    txtFile(0).Enabled = bState
    txtFile(1).Enabled = bState
    lblFile(0).Enabled = bState
    lblFile(1).Enabled = bState
    optFileFmt(0).Enabled = bState
    optFileFmt(1).Enabled = bState

End Sub

Private Sub cmdFile_Click()

    With moTree
        Select Case True
            Case (optFileOp(efoLoad).Value = True)
                If pValidateFile Then
                    .Load msFile, Abs(CLng(optFileFmt(1).Value = True))
                End If
            Case (optFileOp(efoSave).Value = True)
                If pValidateFile Then
                    .Save msFile, Abs(CLng(optFileFmt(1).Value = True))
                End If
            Case Else
                .ClearTreeView
        End Select
    End With

End Sub

Private Sub pInitFileOp()
    If msFile = "" Then
        txtFile(0).Text = "Tut1Tree.dat"
        txtFile(1).Text = App.Path
    End If
End Sub

Private Function pValidateFile() As Boolean
    msFile = App.Path + "\" + txtFile(0).Text
    Select Case True
        Case optFileOp(efoLoad).Value = True
            If Not pFileExist(msFile) Then
                MsgBox msFile + " does not exist", vbApplicationModal + vbInformation, "File Operation"
                msFile = ""
                txtFile(0).SetFocus
            Else
                pValidateFile = True
            End If

        Case optFileOp(efoSave).Value = True
            If pFileExist(msFile) Then
                If MsgBox(msFile + " already exists - Overwrite", _
                          vbApplicationModal + vbInformation + vbYesNo + vbDefaultButton2, _
                          "File Operation") = vbNo Then
                    msFile = ""
                    txtFile(0).SetFocus
                Else
                    VBA.Kill msFile
                    pValidateFile = True
                End If
            Else
                pValidateFile = True
            End If

    End Select
End Function

'===========================================================================
' TAB 4 : Scroll
'
Private Sub cmdRemote_Click(Index As Integer)
    '
    '## Progmatically control the TreeView scrolling and Expanded/Collapsed
    '   node states
    '
    Select Case Index
        Case [Top] To [Right]
                moTree.ScrollView CByte(Index)

        Case [Reset]
                With moTree
                    .ScrollView [Left]
                    .ScrollView [Top]
                End With

        Case [Expand Selected]
                With tvwDialog
                    If .SelectedItem.Children Then
                        If Not .SelectedItem.Expanded Then
                            .Visible = False
                            .SelectedItem.Expanded = True
                            .SelectedItem.EnsureVisible
                            .Visible = True
                        Else
                            MsgBox "Selected Node already expanded!", vbExclamation, "WARNING!"
                        End If
                    Else
                        MsgBox "Selected Node has no children!", vbExclamation, "WARNING!"
                    End If
                End With

        Case [Expand]
                With tvwDialog
                    If .SelectedItem.Children Then
                        If Not .SelectedItem.Expanded Then
                            .Visible = False
                            moTree.ExpandChildNodes .SelectedItem
                            .SelectedItem.EnsureVisible
                            .Visible = True
                        Else
                            MsgBox "Selected Node already expanded!", vbExclamation, "WARNING!"
                        End If
                    Else
                        MsgBox "Selected Node has no children!", vbExclamation, "WARNING!"
                    End If
                End With

        Case [Expand All]
                With tvwDialog
                    .Visible = False
                    moTree.ExpandAll
                    .SelectedItem.EnsureVisible
                    .Visible = True
                End With

        Case [Collape Selected]
                With tvwDialog
                    If .SelectedItem.Children Then
                        If .SelectedItem.Expanded Then
                            .SelectedItem.Expanded = False
                        Else
                            MsgBox "Selected Node already collapsed!", vbExclamation, "WARNING!"
                        End If
                    Else
                        MsgBox "Selected Node has no children!", vbExclamation, "WARNING!"
                    End If
                End With

        Case [Collape]
                With tvwDialog
                    If .SelectedItem.Children Then
                        If .SelectedItem.Expanded Then
                            moTree.CollapseChildNodes .SelectedItem
                        Else
                            MsgBox "Selected Node already collapsed!", vbExclamation, "WARNING!"
                        End If
                    Else
                        MsgBox "Selected Node has no children!", vbExclamation, "WARNING!"
                    End If
                End With

        Case [Collapse All]
                With tvwDialog
                    .Visible = False
                    moTree.CollapseAll
                    .Visible = True
                End With

    End Select
End Sub

'===========================================================================
' TAB 5 : Move
'
Private Sub cmdMove_Click(Index As Integer)

    Dim oNode   As MSComctlLib.Node
    Dim oMarker As MSComctlLib.Node
    Dim eRelationship As TreeRelationshipConstants

    Set oNode = tvwDialog.SelectedItem
    If (oNode Is Nothing) Then Exit Sub

    Select Case Index
        Case [eUp]
            If oNode = oNode.FirstSibling Then Exit Sub
            Set oMarker = oNode.Previous
            eRelationship = tvwPrevious

        Case [eDown]
            If oNode = oNode.LastSibling Then Exit Sub
            Set oMarker = oNode.Next
            eRelationship = tvwNext

        Case [eFirst]
            If oNode = oNode.FirstSibling Then Exit Sub
            Set oMarker = oNode.FirstSibling
            eRelationship = tvwFirst

        Case [eLast]
            If oNode = oNode.LastSibling Then Exit Sub
            Set oMarker = oNode.LastSibling
            eRelationship = tvwLast

    End Select
    
    If Not moTree.NodeMove(oMarker, oNode, , eRelationship) Then
        '## Should never get here!!
        MsgBox "Unable to move node '" + oNode.Text + "'.", vbExclamation, "Move Node"
    End If
    tvwDialog.SetFocus

End Sub

'===========================================================================
' Internal Functions
'
Private Function pFileExist(FileName As String) As Boolean

   Dim TempAttr As Integer
   
   On Error GoTo ErrorFileExist 'any errors show that the file doesnt exist, so goto this label

   TempAttr = GetAttr(FileName) 'get the attributes of the files
   pFileExist = ((TempAttr And vbDirectory) = 0) 'check if its a directory and not a file
   GoTo ExitFileExist
   
ErrorFileExist:
   pFileExist = False 'return that the file doesnt exist
   Resume ExitFileExist 'carry on with the code
   
ExitFileExist:
   On Error GoTo 0 'clear all errors
   
End Function

Private Sub pHiLite(txtBox As TextBox)
    '
    '## Selects all text of the designated Textbox
    '
    With txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub pInitData()                                                     '@@ v01.01.00

#If OLDDATA = 1 Then
    With moTree
        .NodeAdd , , "A", "Basic Functions", 1, 3, , , , True, , True, , , , 2
        .NodeAdd , , "B", "Drag and Drop", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "C", "MultiSelection", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "D", "Load On Demand", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "E", "ADO Integration", 1, 3, , , , , , , , , , 2

        Dim lLoop As Long
        .NodeAdd , , "X1", "Node Item 1", 1, 3, , , , , , , , , , 2
        For lLoop = 2 To 50
            .NodeAdd tvwDialog.Nodes("X" + CStr(lLoop - 1)), tvwChild, "X" + CStr(lLoop), "Node Item " + CStr(lLoop), 1, 3, , , , , , , , , , 2
        Next
    End With
#Else
    Dim lLoop1 As Long
    Dim lLoop2 As Long
    Dim lLoop3 As Long
    Dim lLoop4 As Long
    Dim sChar1 As String * 1
    Dim sChar2 As String * 2
    Dim sChar3 As String * 3
    Dim sText  As String

    With moTree
        For lLoop1 = 65 To 70
            sChar1 = Chr$(lLoop1)
            .NodeAdd , , sChar1, "Folder " + sChar1, 5, 6, , , , , , , , , , 2
            For lLoop2 = 1 To 5
                sChar2 = sChar1 + CStr(lLoop2)
                .NodeAdd tvwDialog.Nodes(sChar1), tvwChild, sChar2, "Sub Folder " + sChar2, 1, 3, , , , , , , , &HFF&, , 2
                For lLoop3 = 1 To 3
                    sChar3 = sChar2 + CStr(lLoop3)
                    sText = "Book ID " + CStr(lLoop3)
                    .NodeAdd tvwDialog.Nodes(sChar2), tvwChild, sChar3, sText, 10, 8, , , , , , , , &HFF0000, , 10
                    For lLoop4 = 1 To 3
                        .NodeAdd tvwDialog.Nodes(sChar3), tvwChild, sChar3 + "-" + CStr(lLoop4), "Chapter " + CStr(lLoop4) + "[" + sText + "]", 11, 9, , , , , , , , &H800080, , 11
                    Next
                Next
            Next
        Next
    End With
#End If

End Sub
