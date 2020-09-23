VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fAniBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demonstration: Animated Scrolling TreeView Graphic Background - Unique!"
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
   Icon            =   "fAniBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTreeHost 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   105
      ScaleHeight     =   6735
      ScaleWidth      =   3585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   3585
      Begin MSComctlLib.TreeView tvwDialog 
         Height          =   6630
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   11695
         _Version        =   393217
         Style           =   7
         Appearance      =   1
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
   Begin VB.Frame Frame1 
      Caption         =   "Animation Settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   4305
      TabIndex        =   2
      Top             =   105
      Width           =   4425
      Begin VB.TextBox txtMotion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1050
         Width           =   855
      End
      Begin VB.VScrollBar vsMotion 
         Height          =   300
         Left            =   3165
         Max             =   100
         Min             =   1
         SmallChange     =   5
         TabIndex        =   8
         Top             =   1050
         Value           =   1
         Width           =   225
      End
      Begin VB.CheckBox chkState 
         Caption         =   "InActive"
         Height          =   330
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.ComboBox cboDirection 
         Height          =   315
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1575
         Width           =   1065
      End
      Begin VB.VScrollBar vsSpeed 
         Height          =   300
         LargeChange     =   100
         Left            =   3165
         Max             =   5000
         SmallChange     =   10
         TabIndex        =   5
         Top             =   525
         Value           =   50
         Width           =   225
      End
      Begin VB.TextBox txtSpeed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lblSettings 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Motion: "
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   945
         TabIndex        =   6
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label lblSettings 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Direction: "
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   945
         TabIndex        =   9
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label lblSettings 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Delay: "
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   945
         TabIndex        =   3
         Top             =   525
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   1365
      Top             =   3150
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
            Picture         =   "fAniBack.frx":27A2
            Key             =   "Closed1"
            Object.Tag             =   "Closed Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":2D3C
            Key             =   "Open1"
            Object.Tag             =   "Open Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":32D6
            Key             =   "Selected"
            Object.Tag             =   "Selected Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":3870
            Key             =   "Group"
            Object.Tag             =   "Group Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":3E0A
            Key             =   "Closed2"
            Object.Tag             =   "Closed Network Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":43A4
            Key             =   "Open2"
            Object.Tag             =   "Open Network Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":493E
            Key             =   "Clock"
            Object.Tag             =   "Clock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":4D90
            Key             =   "Barcode"
            Object.Tag             =   "Barcode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":4EEA
            Key             =   "Agent"
            Object.Tag             =   "Agent"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":7024
            Key             =   "Diary"
            Object.Tag             =   "Diary"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":CC46
            Key             =   "Item"
            Object.Tag             =   "Card Item"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":CF60
            Key             =   "ShareOverlay"
            Object.Tag             =   "Share Overlay"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":D4FA
            Key             =   "ShortcutOverlay"
            Object.Tag             =   "Shortcut Overlay"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":D654
            Key             =   "Custom1Overlay"
            Object.Tag             =   "Custom Overlay 1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":D96E
            Key             =   "Custom2Overlay"
            Object.Tag             =   "Custom Overlay 2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAniBack.frx":DAC8
            Key             =   "Custom3Overlay"
            Object.Tag             =   "Custom Overlay 3"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6690
      Left            =   105
      Picture         =   "fAniBack.frx":DC22
      ScaleHeight     =   6690
      ScaleWidth      =   2715
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Click to view original size"
      Top             =   105
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox picAnimate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   735
      ScaleHeight     =   6735
      ScaleWidth      =   2715
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblComments 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comments:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6090
      TabIndex        =   12
      Top             =   3255
      Width           =   855
   End
End
Attribute VB_Name = "fAniBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fAniBack [Demo 1]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         16/08/2002
' Version:      01.00.00
' Description:  TreeView background animation demonstration.
' Edit History: 01.00.00 16/08/2002 Initial Release
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

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Const HWND_TOPMOST   As Long = -1
    Private Const SWP_NOMOVE     As Long = 2
    Private Const SWP_NOSIZE     As Long = 1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const WM_SETREDRAW   As Long = &HB
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private mlScanline As Long
Private mlMotion   As Long
Private mlHeight   As Long
Private mlWidth    As Long

Private WithEvents moTimer As XTimer
Attribute moTimer.VB_VarHelpID = -1

Private Sub Form_Load()

#If NODLL = 0 Then
    Set moTree = New vbTree.cTreeView   '## Used to manage the TreeView
#Else
    Set moTree = New cTreeView
#End If

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog
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
        '## Initialise timer & image
        '
        Set moTimer = New XTimer

        With picBackground
            mlHeight = .Height \ Screen.TwipsPerPixelY
            mlWidth = .Width \ Screen.TwipsPerPixelX
        End With
        mlScanline = mlHeight
        vsMotion_Change
        vsSpeed_Change
        chkState_Click
        .BackMode = bmGraphic
        pUpdatePic
        Set .BackPicture = picAnimate.Picture
        '
        '## Enable the treeview to display the nodes
        '
        .Redraw True
    End With
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Unload(Cancel As Integer)
    moTimer.Enabled = False
    Set moTimer = Nothing
    Set moTree = Nothing
End Sub

Private Sub chkState_Click()
    moTimer.Enabled = (chkState.Value = vbChecked)
    chkState.Caption = IIf(chkState.Value = vbChecked, "Active", "InActive")
End Sub

Private Sub moTimer_Tick()

    Static InProc As Boolean

    If InProc Then Exit Sub
    InProc = True
    pUpdatePic
    pLockRefresh picTreeHost.hWnd, True
    Set moTree.BackPicture = picAnimate.Picture
    pLockRefresh picTreeHost.hWnd, False
    InvalidateRect picTreeHost.hWnd, ByVal 0&, 0
    InvalidateRect tvwDialog.hWnd, ByVal 0&, 0
    DoEvents
    InProc = False

End Sub

Private Sub vsMotion_Change()
    txtMotion.Text = CStr(vsMotion.Value) + " pixels"
    mlMotion = vsMotion.Value
    SendKeys "{Tab}"
End Sub

Private Sub vsSpeed_Change()
    txtSpeed.Text = CStr(vsSpeed.Value) + " ms"
    moTimer.Interval = vsSpeed.Value
    SendKeys "{Tab}"
End Sub

Private Sub pInitData()

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

    With cboDirection
        .AddItem "Down": .ItemData(.NewIndex) = -1
        .AddItem "Up":   .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    lblComments.Caption = vbCrLf + vbCrLf + _
                          "  This code was an experiment to see  " + vbCrLf + _
                          "  what was possible with the background image  " + vbCrLf + _
                          "  and a way to overcome the initial flickering  " + vbCrLf + _
                          "  encountered during the development of this demo.  " + vbCrLf + _
                          "  The demo was also put together in response to  " + vbCrLf + _
                          "  those who have pushed me to add graphic  " + vbCrLf + _
                          "  background to the comtrol. Well, here it is  " + vbCrLf + _
                          "  guys - enjoy!  " + vbCrLf + vbCrLf + _
                          "  If you like what's been done (pretty unique hey!),  " + vbCrLf + _
                          "  then please vote for the wrapper - Thanx!" + vbCrLf + vbCrLf

End Sub

Private Function pLockRefresh(hWnd As Long, cLock As Boolean)
    '
    '## Enable/disable an object from painting... Note: This gives more
    '   control than LockWindowUpdate API - so be careful if you use this!
    '
    SendMessage hWnd, WM_SETREDRAW, Not cLock, 0
End Function

Private Sub pUpdatePic()

    mlScanline = mlScanline + mlMotion * cboDirection.ItemData(cboDirection.ListIndex)
    If mlScanline > mlHeight Then mlScanline = 1
    If mlScanline < 1 Then mlScanline = mlHeight
    BitBlt picAnimate.hDC, 0, 0, mlWidth, mlHeight - mlScanline, _
           picBackground.hDC, 0, mlScanline, vbSrcCopy
    BitBlt picAnimate.hDC, 0, mlHeight - mlScanline + 1, mlWidth, mlScanline, _
           picBackground.hDC, 0, 0, vbSrcCopy
    Set picAnimate.Picture = picAnimate.Image

End Sub
