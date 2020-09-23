VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fMultiSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 8 : Disabled Multi-Node Selection"
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
   Icon            =   "fMultiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDialog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "R&eset Disabled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2865
      Width           =   1215
   End
   Begin VB.ListBox lstDialog 
      BackColor       =   &H80000018&
      Height          =   3570
      Index           =   1
      ItemData        =   "fMultiSelect.frx":27A2
      Left            =   9345
      List            =   "fMultiSelect.frx":27A4
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3990
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.ListBox lstDialog 
      BackColor       =   &H80000018&
      Height          =   3570
      Index           =   0
      ItemData        =   "fMultiSelect.frx":27A6
      Left            =   9345
      List            =   "fMultiSelect.frx":27A8
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   3150
   End
   Begin MSComctlLib.TabStrip tabDialog 
      Height          =   4110
      Index           =   0
      Left            =   5880
      TabIndex        =   13
      Top             =   105
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7250
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Multi-Select Nodes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Disabled Nodes"
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
   Begin VB.CommandButton cmdDialog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Disable Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   4305
      TabIndex        =   10
      Top             =   4305
      Width           =   4950
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   1380
         Index           =   1
         Left            =   3045
         ScaleHeight     =   1380
         ScaleWidth      =   3795
         TabIndex        =   20
         Top             =   1365
         Visible         =   0   'False
         Width           =   3795
         Begin VB.CheckBox chkOptDisabled 
            Appearance      =   0  'Flat
            Caption         =   "Allow selection of disabled nodes"
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
            Height          =   225
            Index           =   2
            Left            =   840
            TabIndex        =   23
            Top             =   810
            Width           =   2805
         End
         Begin VB.CheckBox chkOptDisabled 
            Appearance      =   0  'Flat
            Caption         =   "Bold Disabled Nodes"
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
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   22
            Top             =   510
            Width           =   2805
         End
         Begin VB.CheckBox chkOptDisabled 
            Appearance      =   0  'Flat
            Caption         =   "Use Default Disabled Colours"
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
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   21
            Top             =   210
            Value           =   1  'Checked
            Width           =   2805
         End
      End
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   1380
         Index           =   0
         Left            =   315
         ScaleHeight     =   1380
         ScaleWidth      =   3795
         TabIndex        =   15
         Top             =   525
         Visible         =   0   'False
         Width           =   3795
         Begin VB.CheckBox chkOptMulti 
            Appearance      =   0  'Flat
            Caption         =   "No Default Selection"
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
            Height          =   225
            Index           =   3
            Left            =   840
            TabIndex        =   19
            Top             =   1110
            Value           =   1  'Checked
            Width           =   2805
         End
         Begin VB.CheckBox chkOptMulti 
            Appearance      =   0  'Flat
            Caption         =   "No Clear On Space Click"
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
            Height          =   225
            Index           =   2
            Left            =   840
            TabIndex        =   18
            Top             =   810
            Width           =   2805
         End
         Begin VB.CheckBox chkOptMulti 
            Appearance      =   0  'Flat
            Caption         =   "Bold Selected Nodes"
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
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   17
            Top             =   525
            Width           =   2805
         End
         Begin VB.CheckBox chkOptMulti 
            Appearance      =   0  'Flat
            Caption         =   "Use Default selection Colours"
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
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   16
            Top             =   210
            Value           =   1  'Checked
            Width           =   2805
         End
      End
      Begin MSComctlLib.TabStrip tabDialog 
         Height          =   2010
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   3545
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "MultiSelect"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Disabled"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "T&ransfer ->>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   4515
      TabIndex        =   7
      Top             =   3465
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Toggle Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4515
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Clear Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   4515
      TabIndex        =   1
      Top             =   315
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Select Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   4515
      TabIndex        =   5
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Select &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   4515
      TabIndex        =   6
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Clear A&ll"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   4515
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "To&ggle All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   4515
      TabIndex        =   4
      Top             =   1395
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoList 
      Appearance      =   0  'Flat
      Caption         =   "Auto"
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
      Height          =   255
      Left            =   4935
      TabIndex        =   8
      Top             =   3885
      Value           =   1  'Checked
      Width           =   645
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
      Left            =   4830
      Top             =   2835
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
            Picture         =   "fMultiSelect.frx":27AA
            Key             =   "Closed1"
            Object.Tag             =   "Closed Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":2D44
            Key             =   "Open1"
            Object.Tag             =   "Open Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":32DE
            Key             =   "Selected"
            Object.Tag             =   "Selected Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":3878
            Key             =   "Group"
            Object.Tag             =   "Group Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":3E12
            Key             =   "Closed2"
            Object.Tag             =   "Closed Network Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":43AC
            Key             =   "Open2"
            Object.Tag             =   "Open Network Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":4946
            Key             =   "Clock"
            Object.Tag             =   "Clock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":4D98
            Key             =   "Barcode"
            Object.Tag             =   "Barcode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":4EF2
            Key             =   "Agent"
            Object.Tag             =   "Agent"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":702C
            Key             =   "Diary"
            Object.Tag             =   "Diary"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":CC4E
            Key             =   "Item"
            Object.Tag             =   "Card Item"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":CF68
            Key             =   "ShareOverlay"
            Object.Tag             =   "Share Overlay"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":D502
            Key             =   "ShortcutOverlay"
            Object.Tag             =   "Shortcut Overlay"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":D65C
            Key             =   "Custom1Overlay"
            Object.Tag             =   "Custom Overlay 1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":D976
            Key             =   "Custom2Overlay"
            Object.Tag             =   "Custom Overlay 2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":DAD0
            Key             =   "Custom3Overlay"
            Object.Tag             =   "Custom Overlay 3"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fMultiSelect [Tutorial 8]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         10/07/2002
' Version:      01.00.00
' Description:  Test/Demo TreeView Handler with Multi-Selection & Diabled Nodes
' Edit History: 01.00.00 10/07/2002 Initial Release
'
' Notes:        To view all GUI features, set form Height = 18000,
'               Width = 21500; to hide, set form Height = 7260,
'               Width = 9450
'
'===========================================================================

Option Explicit

Private Const cSELFORECOLOR As Long = vbYellow
Private Const cSELBACKCOLOR As Long = vbRed

Private Const cDisFORECOLOR As Long = vbHighlightText
Private Const cDisBACKCOLOR As Long = vb3DLight

#Const OLDDATA = 0      '## 1 = Old Test Data

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As vbTree.iMultiSelect
    Private miDisabled        As vbTree.iDisable
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As iMultiSelect
    Private miDisabled        As iDisable
#End If

Private Enum eCommand
    eClear = 0
    eClearAll = 1
    eToggle = 2
    eToggleAll = 3
    eSelect = 4
    eSelectAll = 5
    eTransfer = 6
    eDisabled = 7
    eResetDisabled = 8
End Enum

Private Enum eCheck
    eDefColor = 0
    eSelBold = 1
    eDisBold = 1
    eNoCLear = 2
    eAllowDisSelect = 2
    eDefSel = 3
End Enum

Private Enum eTabStrip
    eResults = 0
    eOptions = 1
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
    Set moTree = New cTreeView          '## Used to manage the TreeView
#End If

    Set miMultiSelect = moTree
    Set miDisabled = moTree

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
            With .Nodes(1)
                .Selected = True
                .Selected = False
            End With
        End With
    End With
    '
    '## Set the initial value/states
    '
    tabDialog(0).SelectedItem = tabDialog(0).Tabs(1)  '## Force correct Frame to be displayed
    tabDialog(1).SelectedItem = tabDialog(1).Tabs(1)  '## Force correct Frame to be displayed
    miDisabled.SelectMode = [eiDis Select Direction]

End Sub

'===========================================================================
' Form Control Events
'
Private Sub chkOptDisabled_Click(Index As Integer)

    Dim sType As String

    With miDisabled
        Select Case Index
            Case eDefColor
                Select Case chkOptDisabled(Index).Value
                    Case vbUnchecked
                        .BackColor = cDisBACKCOLOR
                        .ForeColor = cDisFORECOLOR
                    Case vbChecked
                        .BackColor = vbWindowBackground
                        .ForeColor = vbGrayText
                End Select

            Case eDisBold
                .Bold = CBool(chkOptDisabled(Index).Value)

            Case eAllowDisSelect
                .AllowSelect = CBool(chkOptDisabled(Index).Value)

        End Select
    End With
    On Error Resume Next    '## Stop error when called from Form_Load.
    tvwDialog.SetFocus

End Sub

Private Sub chkOptMulti_Click(Index As Integer)

    With miMultiSelect
        Select Case Index
            Case eDefColor
                Select Case chkOptMulti(Index).Value
                    Case vbUnchecked
                        .SelBackColor = cSELBACKCOLOR
                        .SelForeColor = cSELFORECOLOR
                    Case vbChecked
                        .SelBackColor = vbHighlight
                        .SelForeColor = vbHighlightText
                End Select

            Case eSelBold
                .SelBold = CBool(chkOptMulti(Index).Value)
            Case eNoCLear
                .NoClearOnSpaceClick = CBool(chkOptMulti(Index).Value)

            Case eDefSel
                .NoDefaultSel = CBool(chkOptMulti(Index).Value)

        End Select
    End With
    On Error Resume Next    '## Stop error when called from Form_Load.
    tvwDialog.SetFocus

End Sub

Private Sub cmdDialog_Click(Index As Integer)

    Dim oNode As MSComctlLib.Node

    Select Case Index
        Case eClear
            With miMultiSelect
                .ClearSelection .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eClearAll
            With miMultiSelect
                .ClearSelection
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eToggle
            With miMultiSelect
                .ToggleSelection .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eToggleAll
            With miMultiSelect
                .ToggleSelection
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eSelect
            With miMultiSelect
                .SelectAllNodes .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eSelectAll
            With miMultiSelect
                .SelectAllNodes
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eDisabled
            miDisabled.DisableMultiSelect

        Case eResetDisabled
            miDisabled.DisableAllNodes , False
            lstDialog(1).Clear

        Case eTransfer
            Dim lNdx As Long
            lNdx = tabDialog(0).SelectedItem.Index - 1
            With lstDialog(lNdx)
                .Visible = False
                .Clear
                Select Case lNdx
                    Case 0  '## Multiselect nodes
                        For Each oNode In miMultiSelect
                            .AddItem oNode.Text
                        Next
                    Case 1  '## Disabled Nodes
                        For Each oNode In miDisabled
                            .AddItem oNode.Text
                        Next
                End Select
                .Visible = True
            End With

    End Select
    On Error Resume Next
    tvwDialog.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

Private Sub tabDialog_Click(Index As Integer)

    Dim lPtr  As Long
    Dim lLoop As Long

    With tabDialog(Index)
        For lLoop = 1 To .Tabs.Count
            lPtr = lLoop - 1
            '
            '## Position and show/hide the container controls
            '
            Select Case Index
                Case eResults
                    lstDialog(lPtr).Visible = (lLoop = .SelectedItem.Index)
                    lstDialog(lPtr).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
                    lstDialog(lPtr).ZOrder                          '## Bring to front
                    '
                    '## Fix tab height to cleanly wrap the List control
                    '
                    .Height = .Height + (lstDialog(lPtr).Height - .ClientHeight)
                Case eOptions
                    picOptions(lPtr).Visible = (lLoop = .SelectedItem.Index)
                    picOptions(lPtr).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
                    picOptions(lPtr).ZOrder                          '## Bring to front
            End Select
        Next
    End With
    If chkAutoList.Value Then
        cmdDialog_Click eTransfer
    End If

End Sub

'===========================================================================
' cTreeView wrapper Events
'
Private Sub moTree_NodeClick(ByVal Node As MSComctlLib.Node)
    With miMultiSelect
        pCmdsEnable Not (.FocusNode Is Nothing)
    End With
End Sub

Private Sub moTree_SelChange()
    If chkAutoList.Value Then
        cmdDialog_Click eTransfer
    End If
End Sub

'===========================================================================
' Internal Functions
'
Private Sub pCmdsEnable(Mode As Boolean)
    cmdDialog(eClear).Enabled = Mode
    cmdDialog(eToggle).Enabled = Mode
    cmdDialog(eSelect).Enabled = Mode
End Sub

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
