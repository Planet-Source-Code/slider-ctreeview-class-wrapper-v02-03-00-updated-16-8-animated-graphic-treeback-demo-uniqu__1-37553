VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTreeHotelDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 5 : ADO Load on Demand Only"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "fTreeHotelDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgDialog 
      Left            =   3045
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":0F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":1AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":2044
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTreeHotelDB.frx":2678
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   11695
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "fTreeHotelDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTreeHotel [Tutorial 5 - Load on Demand]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         23/03/2002
' Version:      02.00.00
' Description:  Test/Demo 2 - Hotels by City by Country
' Edit History: 01.00.00 05/12/2001 Initial Release
'               02.00.00 23/03/2002 Modified to work with vbTree
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
Private Const clCITYCOLOR  As Long = &H800080
Private Const clHOTELCOLOR As Long = &H40C0&

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
#End If

Private mcDB               As cDB
Private mbIsBusy           As Boolean

'===========================================================================
' Private: ADO Declarations
'
'## Get Countries
Private Const mcSQL_COUNTRY As String = "SELECT DISTINCTROW Desc, PkID " + _
                                        "FROM [Country] "

'## Get Cities by Country ID
Private Const mcSQL_CITY1   As String = "SELECT DISTINCTROW Desc, PkID " + _
                                        "FROM [City] " + _
                                        "WHERE ((LinkID)="

Private Const mcSQL_CITY2   As String = ") "

'## Get Cities by Country ID
Private Const mcSQL_HOTEL1  As String = "SELECT DISTINCTROW Desc, PkID " + _
                                        "FROM [Hotel] " + _
                                        "WHERE ((LinkID)="

Private Const mcSQL_HOTEL2  As String = ") "

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    Set moTree = New cTreeView

    With moTree
        .HookCtrl tvwDialog
        .Redraw False
        With .Ctrl
            .Style = tvwTreelinesPlusMinusPictureText
            .LineStyle = tvwRootLines
            .Indentation = 10
            .ImageList = imgDialog
            .FullRowSelect = False
            .HideSelection = False
            .HotTracking = True
            .LabelEdit = tvwManual

            pInitData
        End With
        .ContextMenuMode = [After Click]
        .DragEnabled = True
        .Redraw True
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

Private Sub tvwDialog_Expand(ByVal Node As MSComctlLib.Node)
    pExpandNode Node
End Sub

'===========================================================================
' Private subroutines and functions
'

'===========================================================================
' Private ADO subroutines and functions
'
Private Sub pInitData()
    '
    '## Initialise database. NOTE: DATABASE MUST BE IN APP PATH.
    '
    Set mcDB = New cDB
    mcDB.InitDB App.Path + "\TreeHotel.mdb", , , , ejvJet4
    pLoadCountries

End Sub

Private Function pExpandNode(Node As MSComctlLib.Node) As Boolean

    Dim TableID As String

    If mbIsBusy Then Exit Function

    With Node
        TableID = VBA.Right$(.Key, 1)
        If InStr("CK", TableID) And Len(.Tag) = 0 Then
            If .Children Then
                tvwDialog.Nodes.Remove .Child.Index
                Select Case TableID
                    Case "C": pLoadCities Node
                    Case "K": pLoadHotels Node
                End Select
                .Tag = "Loaded"
            End If
        End If
    End With

End Function
Private Sub pLoadCountries()

    Dim lLoop  As Long
    Dim lCount As Long
    Dim sKey   As String
    Dim oRs    As ADODB.Recordset
    Dim oNode  As MSComctlLib.Node

    mbIsBusy = True
    '
    '## Did we successfully create the recordset?
    '
    If mcDB.CreateRS(oRs, mcSQL_COUNTRY) Then
        '
        '## Yes.
        '
        lCount = mcDB.RecordCount(oRs)
        If lCount Then
            moTree.Redraw False
            With oRs
                '
                '## For every branch
                '
                For lLoop = 1 To lCount
                    sKey = CStr(!PkID) + "C"
                    moTree.NodeAdd , , sKey, !Desc, 1, 1, , True, , , False, , , , , 2
                    '
                    '## Add dummy key to Group node to display the Expand Handle (plus sign)
                    '
                    moTree.NodeAdd sKey, tvwChild, sKey + "D", "Dummy", , , , , , , False
                    '
                    '## Next node/record
                    '
                    mcDB.MoveDB emrMoveNext, oRs
                Next
            End With
            moTree.Redraw True
        End If
    End If
    mbIsBusy = False

End Sub

Private Sub pLoadCities(oNode As Node)

    Dim lLoop  As Long
    Dim lCount As Long
    Dim lPtr   As Long
    Dim sKey   As String
    Dim oRs    As ADODB.Recordset

    With oNode
        Debug.Print "** Loading Cities for Node [" + oNode.Text + "]"
        If mcDB.CreateRS(oRs, mcSQL_CITY1 + CStr(Val(.Key)) + mcSQL_CITY2) Then
            lCount = mcDB.RecordCount(oRs)
            If lCount Then
                For lLoop = 1 To lCount
                    '
                    '## Add product nodes to group node
                    '@@ v 01.00.01 Added icon pointers
                    '
                    sKey = CStr(oRs!PkID) + "K"
                    moTree.NodeAdd .Key, tvwChild, sKey, oRs!Desc, 3, 3, , , , , False, , , clCITYCOLOR, , 4
                    '
                    '## Add dummy key to Group node to display the Expand Handle (plus sign)
                    '
                    moTree.NodeAdd sKey, tvwChild, sKey + "D", "Dummy", , , , , , , False
                    '
                    '## Next node/record
                    '
                    mcDB.MoveDB emrMoveNext, oRs
                Next
            End If
        End If
    End With

End Sub

Private Sub pLoadHotels(oNode As Node)

    Dim lLoop  As Long
    Dim lCount As Long
    Dim lPtr   As Long
    Dim oRs    As ADODB.Recordset

    With oNode
        Debug.Print "** Loading Hotels for Node [" + oNode.Text + "]"
        If mcDB.CreateRS(oRs, mcSQL_HOTEL1 + CStr(Val(.Key)) + mcSQL_HOTEL2) Then
            lCount = mcDB.RecordCount(oRs)
            If lCount Then
                For lLoop = 1 To lCount
                    '
                    '## Add product nodes to group node
                    '@@ v 01.00.01 Added icon pointers
                    '
                    moTree.NodeAdd .Key, tvwChild, CStr(oRs!PkID) + "H", oRs!Desc, 6, 6, , , , , , , , clHOTELCOLOR, , 6
                    '
                    '## Next node/record
                    '
                    mcDB.MoveDB emrMoveNext, oRs
                Next
            End If
        End If
    End With

End Sub
