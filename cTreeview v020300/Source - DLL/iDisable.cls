VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iDisable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Interface:   iDisable
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        10/07/2002
' Version:     01.01.00     (v02.03.00 cTreeView)
' Description: Interface definition for Disable Node feature
' Edit History: 01.00.00 23/03/2002 Initial Release
'               01.01.00 10/07/2002 Added support for Multi-Select Nodes
' Notes:       Due to the unique nature of Disable Node, the
'              interface clearly defined the properties and methods used
'              for this feature.
' Compiling:   DLL : set class Instancing = 2 - PublicNotCreatable
'              APP : set class Instancing = 1 - Private
'
'===========================================================================

Option Explicit

Public Enum eDisableAdjustSelectMode
    [eiDis Select Direction] = 0        '##
    [eiDis Focus Rectangle] = 1         '## (Default)
    [eiDis Select None] = 2             '##
End Enum

Public Property Get AllowSelect() As Boolean

End Property

Public Property Let AllowSelect(ByVal New_State As Boolean)

End Property

Public Property Get BackColor() As OLE_COLOR

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

End Property

Public Property Get Bold() As Boolean

End Property

Public Property Let Bold(ByVal New_Bold As Boolean)

End Property

Public Property Get DisabledNode(ByVal Index As Long) As Collection

End Property

Public Property Get FocusNode() As MSComctlLib.Node

End Property

Public Property Get ForeColor() As OLE_COLOR

End Property

Public Property Let ForeColor(ByVal New_SelForeColor As OLE_COLOR)

End Property

Public Property Get SelectMode() As eDisableAdjustSelectMode

End Property

Public Property Let SelectMode(ByVal New_Mode As eDisableAdjustSelectMode)

End Property

Public Sub DisableAllNodes(Optional ByRef Node As MSComctlLib.Node, _
                           Optional ByVal State As Boolean = True)

End Sub

Public Sub DisableNode(ByRef Node As MSComctlLib.Node, _
              Optional ByVal State As Boolean = True)

End Sub

Public Sub DisableMultiSelect(Optional ByVal State As Boolean = True)   '@@ v01.01.00 (02.03.00)

End Sub

Public Sub ToggleDisabled(Optional ByRef Node As MSComctlLib.Node)

End Sub

Public Sub ToggleNodeState(Optional ByRef Node As MSComctlLib.Node, _
                           Optional ByVal Children As Boolean)

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

End Function
