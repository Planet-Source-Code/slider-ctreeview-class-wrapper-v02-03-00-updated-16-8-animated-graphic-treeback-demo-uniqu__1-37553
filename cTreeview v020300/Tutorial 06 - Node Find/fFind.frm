VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 6 : Find a Node or Multiple Nodes Using 3 Different Searching Methods"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDialog 
      Appearance      =   0  'Flat
      Caption         =   "Select all search results in TreeView (Multi-Node Selection)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4305
      TabIndex        =   34
      Top             =   6510
      Width           =   5055
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Find a group of Nodes (Pattern matching...) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Index           =   2
      Left            =   9450
      TabIndex        =   19
      Top             =   6930
      Visible         =   0   'False
      Width           =   4845
      Begin VB.Frame fraFind3 
         Caption         =   "Compare: "
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
         Index           =   1
         Left            =   210
         TabIndex        =   24
         ToolTipText     =   "Compare Method"
         Top             =   1260
         Width           =   1275
         Begin VB.OptionButton optFind3Comp 
            Appearance      =   0  'Flat
            Caption         =   "Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   210
            TabIndex        =   26
            ToolTipText     =   "Sort Order = (A=a) < (À=à) < (B=b) < (E=e) < (Ê=ê) < (Z=z) < (Ø=ø)"
            Top             =   525
            Width           =   855
         End
         Begin VB.OptionButton optFind3Comp 
            Appearance      =   0  'Flat
            Caption         =   "Binary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   25
            ToolTipText     =   "Sort Order = A < B < E < Z < a < b < e < z < À < Ê < Ø < à < ê < ø"
            Top             =   210
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtFind3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1575
         TabIndex        =   23
         ToolTipText     =   "Click on a Node in the TreeView Control or leave blank to search all nodes"
         Top             =   840
         Width           =   2745
      End
      Begin VB.Frame fraFind3 
         Caption         =   "Search Type: "
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
         Index           =   0
         Left            =   1575
         TabIndex        =   27
         Top             =   1260
         Width           =   2745
         Begin VB.OptionButton optFind3 
            Appearance      =   0  'Flat
            Caption         =   "Node Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   28
            Top             =   210
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton optFind3 
            Appearance      =   0  'Flat
            Caption         =   "Node Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   210
            TabIndex        =   29
            Top             =   525
            Width           =   1905
         End
         Begin VB.OptionButton optFind3 
            Appearance      =   0  'Flat
            Caption         =   "Node Text and Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   210
            TabIndex        =   30
            Top             =   840
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "GO!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3465
         TabIndex        =   31
         Top             =   2730
         Width           =   1170
      End
      Begin VB.TextBox txtFind3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1575
         TabIndex        =   21
         ToolTipText     =   "Any string expression conforming to the pattern-matching conventions (See Help for more details)"
         Top             =   420
         Width           =   2745
      End
      Begin VB.Label lblFind3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Start Node: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   210
         TabIndex        =   22
         ToolTipText     =   "Click on a Node in the TreeView Control"
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblFind3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   210
         TabIndex        =   20
         ToolTipText     =   "Any string expression conforming to the pattern-matching conventions (See Help for more details)"
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Find a group of Nodes (Containing...) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Index           =   1
      Left            =   9450
      TabIndex        =   10
      Top             =   3465
      Visible         =   0   'False
      Width           =   4845
      Begin VB.CheckBox chkFind2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Case Sensitive Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   1890
         TabIndex        =   13
         Top             =   840
         Width           =   2430
      End
      Begin VB.Frame fraFind2 
         Caption         =   "Search Type: "
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
         Left            =   1575
         TabIndex        =   14
         Top             =   1260
         Width           =   2745
         Begin VB.OptionButton optFind2 
            Appearance      =   0  'Flat
            Caption         =   "Node Text and Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   210
            TabIndex        =   17
            Top             =   840
            Width           =   1905
         End
         Begin VB.OptionButton optFind2 
            Appearance      =   0  'Flat
            Caption         =   "Node Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   210
            TabIndex        =   16
            Top             =   525
            Width           =   1905
         End
         Begin VB.OptionButton optFind2 
            Appearance      =   0  'Flat
            Caption         =   "Node Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   15
            Top             =   210
            Value           =   -1  'True
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "GO!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3465
         TabIndex        =   18
         Top             =   2730
         Width           =   1170
      End
      Begin VB.TextBox txtFind2 
         Appearance      =   0  'Flat
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
         Left            =   1575
         TabIndex        =   12
         Top             =   420
         Width           =   2745
      End
      Begin VB.Label lblFind2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Find a specific Node (Exact Match Case-Insensitive) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Index           =   0
      Left            =   9450
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   4845
      Begin VB.TextBox txtFind1 
         Appearance      =   0  'Flat
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
         Left            =   1575
         TabIndex        =   4
         Top             =   420
         Width           =   2745
      End
      Begin VB.CheckBox chkFind1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Match Node Key"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   1890
         TabIndex        =   5
         Top             =   840
         Value           =   1  'Checked
         Width           =   2430
      End
      Begin VB.TextBox txtFind1 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         Top             =   1260
         Width           =   2745
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "GO!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3465
         TabIndex        =   9
         Top             =   2730
         Width           =   1170
      End
      Begin VB.CheckBox chkFind1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Select Node on Success"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1890
         TabIndex        =   8
         Top             =   1680
         Width           =   2430
      End
      Begin VB.Label lblFind1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Node Text: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label lblFind1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Node Key: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   1260
         Width           =   1275
      End
   End
   Begin MSComctlLib.TabStrip tabDialog 
      Height          =   3690
      Left            =   4305
      TabIndex        =   1
      Top             =   105
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   6509
      TabWidthStyle   =   1
      Style           =   2
      TabFixedWidth   =   1939
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Single Result"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Simple Multi"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adv. Multi"
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
   Begin VB.ListBox lstDialog 
      BackColor       =   &H80000018&
      Height          =   2400
      Left            =   4305
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Click on a search result to update the TreeView control and display the Selected Node"
      Top             =   4095
      Width           =   4935
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4110
      _ExtentX        =   7250
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
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   4095
      Top             =   4305
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
            Picture         =   "fFind.frx":27A2
            Key             =   "Closed1"
            Object.Tag             =   "Closed Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":2D3C
            Key             =   "Open1"
            Object.Tag             =   "Open Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":32D6
            Key             =   "Selected"
            Object.Tag             =   "Selected Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":3870
            Key             =   "Group"
            Object.Tag             =   "Group Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":3E0A
            Key             =   "Closed2"
            Object.Tag             =   "Closed Network Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":43A4
            Key             =   "Open2"
            Object.Tag             =   "Open Network Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":493E
            Key             =   "Clock"
            Object.Tag             =   "Clock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":4D90
            Key             =   "Barcode"
            Object.Tag             =   "Barcode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":4EEA
            Key             =   "Agent"
            Object.Tag             =   "Agent"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":7024
            Key             =   "Diary"
            Object.Tag             =   "Diary"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":CC46
            Key             =   "Item"
            Object.Tag             =   "Card Item"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":CF60
            Key             =   "ShareOverlay"
            Object.Tag             =   "Share Overlay"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":D4FA
            Key             =   "ShortcutOverlay"
            Object.Tag             =   "Shortcut Overlay"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":D654
            Key             =   "Custom1Overlay"
            Object.Tag             =   "Custom Overlay 1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":D96E
            Key             =   "Custom2Overlay"
            Object.Tag             =   "Custom Overlay 2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fFind.frx":DAC8
            Key             =   "Custom3Overlay"
            Object.Tag             =   "Custom Overlay 3"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEvent 
      AutoSize        =   -1  'True
      Caption         =   "secs."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   8715
      TabIndex        =   38
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Taken:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   6690
      TabIndex        =   36
      Top             =   3780
      Width           =   1275
   End
   Begin VB.Label lblEvent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "0.00000"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   7980
      TabIndex        =   37
      Top             =   3780
      Width           =   750
   End
   Begin VB.Label lblDialog 
      Caption         =   "0 Nodes found."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   5145
      TabIndex        =   35
      Top             =   3780
      Width           =   1275
   End
   Begin VB.Label lblDialog 
      Alignment       =   1  'Right Justify
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   4305
      TabIndex        =   32
      Top             =   3780
      Width           =   750
   End
End
Attribute VB_Name = "fFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fFind [Tutorial 6]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         01/08/2002
' Version:      01.00.00
' Description:  Test/Demo TreeView Handler
' Edit History: 01.00.00 01/08/2002 Initial Release
'
' Notes:        To view all GUI features, set form Height = 18000,
'               Width = 21500; to hide, set form Height = 7260,
'               Width = 9450
'
'===========================================================================

Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Const WM_SETREDRAW        As Long = &HB

#Const OLDDATA = 0      '## 1 = Old Test Data

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As vbTree.iMultiSelect
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As iMultiSelect
#End If

Private meFind2Mode           As eNodeFindEx            '## Used in Find Method #2
Private meFind3Mode           As eNodeFindEx            '## Used in Find Method #3
Private meFind3CompareMode    As VbCompareMethod        '
Private moStartNode           As MSComctlLib.Node       '

Private moTab                 As cTabStrip              '## TabStrip management class

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
    Set moTree = New cTreeView          '## Used to manage the TreeView
#End If

    Set miMultiSelect = moTree

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog, [Multi Select]
        .Appearance = Thin
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
            '
            '## Build TreeView data
            '
            .Visible = False
            pInitData
            .Visible = True
            '
            '## Show focus rectangle over first node but don't select
            '
            .Nodes(1).Selected = True
        End With
    End With
    '
    '## Set the initial value/states
    '
    chkFind1(0).Value = vbUnchecked             '## Force Find1's Node Key match to be hidden
    optFind3Comp(1).Value = True                '## Force Text Compare
    '
    '## Hook the TabStrip control and associate controls with each Tab.
    '
    Set moTab = New cTabStrip
    With moTab
        .HookCtrl tabDialog
        .AutoFit = True
        .Attach 1, fraDialog(0)
        .Attach 2, fraDialog(1)
        .Attach 3, fraDialog(2)
        Set .TabStipContainer = Me
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

'===========================================================================
' Form Control Events
'
Private Sub chkDialog_Click()

    Dim iLoop As Long

    If lstDialog.ListCount Then
        With tvwDialog
            Select Case chkDialog.Value
                Case vbUnchecked
                    miMultiSelect.ClearSelection
                    If lstDialog.ListIndex > -1 Then
                        lstDialog_Click
                    Else
                        Set .SelectedItem = Nothing
                    End If
                Case vbChecked
                    miMultiSelect.ClearSelection
                    For iLoop = 0 To lstDialog.ListCount - 1
                        miMultiSelect.SelectNode .Nodes(lstDialog.ItemData(iLoop)), True
                    Next
                    .SetFocus
            End Select
        End With
    End If

End Sub

Private Sub cmdDialog_Click(Index As Integer)

    Dim oNode  As MSComctlLib.Node
    Dim oNodes As Collection
    Dim oTmr   As cBenchmark

    On Error GoTo ErrorHandler

    lstDialog.Clear
    Set oTmr = New cBenchmark
    With moTree
        Select Case Index
            Case 0  '## ------ Find 1 :: Single Result ------
                oTmr.Start
                If .NodeFind(oNode, _
                             txtFind1(0).Text, _
                             txtFind1(1).Text, _
                             (chkFind1(1).Value = vbChecked)) Then
                    oTmr.Finish
                    pShowTimer oTmr.ElapsedTime
                    '
                    '## Success - Show Results
                    '
                    lstDialog.AddItem oNode.Text + " [Key = '" + oNode.Key + "']"
                    lstDialog.ItemData(0) = oNode.Index
                    lblDialog(1).Caption = "1 Node found."
                Else
                    pShowError Index
                End If

            Case 1  '## ------ Find 2 :: Simple Multi ------
                If Len(Trim$(txtFind2(0).Text)) Then
                    oTmr.Start
                    Set oNodes = .NodeFindEx(txtFind2(0).Text, _
                                             (meFind2Mode + 1), _
                                             (chkFind2(0).Value = vbChecked))
                    oTmr.Finish
                End If
                If Not (oNodes Is Nothing) Then
                    If oNodes.Count > 0 Then
                        '
                        '## Success - Show Results
                        '
                        pShowTimer oTmr.ElapsedTime
                        pShowResults oNodes
                    Else
                        
                        pShowError Index
                    End If
                Else
                    pShowError Index
                End If

            Case 2  '## ------ Find 3 :: Adv. Multi ------
                If Len(Trim$(txtFind3(1).Text)) = 0 Then
                    Set moStartNode = Nothing
                End If
                oTmr.Start
                Set oNodes = .NodeFindEx2(txtFind3(0).Text, _
                                          moStartNode, _
                                          (meFind3Mode + 1), _
                                          meFind3CompareMode)
                oTmr.Finish
                If Not (oNodes Is Nothing) Then
                    If oNodes.Count > 0 Then
                        '
                        '## Success - Show Results
                        '
                        pShowTimer oTmr.ElapsedTime
                        pShowResults oNodes
                    Else
                        pShowError Index
                    End If
                Else
                    pShowError Index
                End If

        End Select
    End With
    Exit Sub

ErrorHandler:
    If Err.Number = 93 Then
        MsgBox Err.Description + ". Please check your pattern matching string for errors.", vbExclamation + vbOKOnly, "Handled Error"
    Else
        MsgBox CStr(Err.Number) + " ::" + Err.Description, vbExclamation + vbOKOnly, "UnHandled Error"
        Unload Me
    End If

End Sub

Private Sub lstDialog_Click()
    With tvwDialog
        If (tabDialog.SelectedItem.Index = 1) Or (chkDialog.Value = vbUnchecked) Then
            miMultiSelect.ClearSelection
            Set .SelectedItem = .Nodes(lstDialog.ItemData(lstDialog.ListIndex))
        Else
            miMultiSelect.FocusNode = .Nodes(lstDialog.ItemData(lstDialog.ListIndex))
            miMultiSelect.FocusNode.EnsureVisible
        End If
        .SetFocus
    End With
End Sub

Private Sub tabDialog_Click()
    '
    '## NOTE: cTabStrip wrapper class manages the TabStrip associated controls!
    '
    If tabDialog.SelectedItem.Index > 1 Then
        chkDialog.Visible = True
        lstDialog.Height = 2400
    Else
        chkDialog.Visible = False
        lstDialog.Height = 2595
    End If

End Sub

'===========================================================================
' Form Control Events :: Find 1
'
Private Sub chkFind1_Click(Index As Integer)

    Dim eState As CheckBoxConstants

    eState = chkFind1(0).Value
    chkFind1(1).Top = IIf(eState = vbUnchecked, 1260, 1680)
    txtFind1(1).Visible = (eState = vbChecked)
    txtFind1(1).Text = ""
    lblFind1(1).Visible = (eState = vbChecked)

End Sub

'===========================================================================
' Form Control Events :: Find 2
'
Private Sub optFind2_Click(Index As Integer)
    meFind2Mode = CByte(Index)
End Sub

'===========================================================================
' Form Control Events :: Find 3
'
Private Sub moTree_Selected(Node As MSComctlLib.Node)
    With tabDialog
        If .SelectedItem.Index = .Tabs.Count Then
            Set moStartNode = Node
            txtFind3(1).Text = moStartNode.Text
        End If
    End With
End Sub

Private Sub optFind3_Click(Index As Integer)
    meFind3Mode = CByte(Index)
End Sub

Private Sub optFind3Comp_Click(Index As Integer)
    meFind3CompareMode = CByte(Index)
End Sub

'===========================================================================
' Internal Functions
'
Private Sub pInitData()

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

Private Function pLockRefresh(ByVal hwnd As Long, ByVal bLock As Boolean)
    '
    '## Enable/disable an object from painting... Note: This gives more
    '   control than LockWindowUpdate API - so be careful if you use this!
    '
    SendMessageLong hwnd, WM_SETREDRAW, Not bLock, 0
End Function

Private Sub pShowError(id As Integer)

    lblDialog(1).Caption = "0 Nodes found."
    Select Case id
        Case 0
            MsgBox "No Node found!" + vbCrLf + _
                   "Make sure that you enter the *full* Node text that you're searching for." + vbCrLf + _
                   "[Example: 'book id 3']", _
                   vbInformation + vbOKOnly, _
                   "Find Node Method 1"
        Case 1
            MsgBox "No Node found!" + vbCrLf + _
                   "Make sure that you enter the *partial* Node text that you're searching for." + vbCrLf + _
                   "[Example: 'id 3']", _
                   vbInformation + vbOKOnly, _
                   "Find Node Method 2"
        Case 2
            MsgBox "No Node found!" + vbCrLf + _
                   "Make sure that you enter the *pattern matching* text that you're searching for." + vbCrLf + _
                   "Note: Text compare method has more success when working with Node Text." + vbCrLf + _
                   "[Example: '*o[l]*2*'] (Note: See Help files for more details)", _
                   vbInformation + vbOKOnly, _
                   "Find Node Method 3"
    End Select

End Sub

Private Sub pShowResults(oNodes As Collection)

    Dim oNode As MSComctlLib.Node

    lblDialog(1).Caption = CStr(oNodes.Count) + " Node" + IIf(oNodes.Count > 1, "s", "") + " found."
    If chkDialog.Value = vbUnchecked Then
        miMultiSelect.ClearSelection
        tvwDialog.SelectedItem.Selected = True
        pLockRefresh lstDialog.hwnd, True
        For Each oNode In oNodes
            lstDialog.AddItem Space$(moTree.NodeNestingLevel(oNode) * 5) + oNode.Text + "    [Key = '" + oNode.Key + "']"
            lstDialog.ItemData(lstDialog.NewIndex) = oNode.Index
        Next
        pLockRefresh lstDialog.hwnd, False
    Else
        miMultiSelect.ClearSelection
        pLockRefresh lstDialog.hwnd, True
        For Each oNode In oNodes
            lstDialog.AddItem Space$(moTree.NodeNestingLevel(oNode) * 5) + oNode.Text + "    [Key = '" + oNode.Key + "']"
            lstDialog.ItemData(lstDialog.NewIndex) = oNode.Index
            miMultiSelect.SelectNode oNode, True
        Next
        pLockRefresh lstDialog.hwnd, False
    End If

End Sub

Private Sub pShowTimer(dElapsed As Double)
    '
    '## Show event to user
    '
    lblEvent(3).Caption = Format$(dElapsed, "#,##0.00000     ")
End Sub
