VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Interface:   iMultiSelect
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        03/02/2001
' Version:     01.00.00
' Description: Interface definition for multi-node selection feature
' Notes:       Due to the unique nature of multi-node selection, the
'              interface clearly defined the properties and methods used
'              for this feature. From a coding perspective, it was a
'              suprise to discover how to expose an enumeration interface
'              using an interface class.
' Compiling:   DLL : set class Instancing = 2 - PublicNotCreatable
'              APP : set class Instancing = 1 - Private
'
'===========================================================================

Option Explicit

Public Property Get NoClearOnSpaceClick() As Boolean
Attribute NoClearOnSpaceClick.VB_Description = "Sets/Returns whether the SpaceBar keyPress event will clear/maintain the current MultiSelect selections."

End Property

Public Property Let NoClearOnSpaceClick(ByVal New_NoClearOnSpaceClick As Boolean)

End Property

Public Property Get NoDefaultSel() As Boolean
Attribute NoDefaultSel.VB_Description = "MultiSelect"

End Property

Public Property Let NoDefaultSel(ByVal New_NoDefaultSel As Boolean)

End Property

Public Property Get SelBackColor() As OLE_COLOR
Attribute SelBackColor.VB_Description = "Sets/Returns custom MultiSelect selection Back Color."

End Property

Public Property Let SelBackColor(ByVal New_SelBackColor As OLE_COLOR)

End Property

Public Property Get SelBold() As Boolean
Attribute SelBold.VB_Description = "Sets/Returns the MultiSelect selection Bold state."

End Property

Public Property Let SelBold(ByVal New_SelBold As Boolean)

End Property

Public Property Get SelForeColor() As OLE_COLOR

End Property

Public Property Let SelForeColor(ByVal New_SelForeColor As OLE_COLOR)

End Property

Public Property Get FocusNode() As MSComctlLib.Node
Attribute FocusNode.VB_Description = "Sets MultiSelect focus node without selecting the node."

End Property

Public Property Let FocusNode(ByVal New_Focus As MSComctlLib.Node)

End Property

'Public Property Get SelInvertColors() As Boolean
'
'End Property
'
'Public Property Let SelInvertColors(ByVal New_SelInvert As Boolean)
''
''To be implemented: Invert colour code:-
''   NewCol = OldCol XOr &HFFFFFF
''
'End Property
'
Public Property Get SelectedNode(ByVal Index As Long) As Collection
Attribute SelectedNode.VB_Description = "Returns a collection of selected Nodes"
Attribute SelectedNode.VB_UserMemId = 0

End Property

Public Sub ClearSelection(Optional ByRef Node As MSComctlLib.Node)
Attribute ClearSelection.VB_Description = "Clears MultiSelect selection."

End Sub

Public Sub SelectAllNodes(Optional ByRef Node As MSComctlLib.Node)
Attribute SelectAllNodes.VB_Description = "Selects all MultiSelect TreeView Nodes."

End Sub

Public Sub SelectNode(ByRef Node As MSComctlLib.Node, _
                      ByVal Selected As Boolean)
Attribute SelectNode.VB_Description = "Sets a MultiSelect Node's select state."

End Sub

Public Sub ToggleNode(Optional ByRef Node As MSComctlLib.Node, _
                      Optional ByVal Children As Boolean)
Attribute ToggleNode.VB_Description = "Toggles a MultiSelect Node and optionally all Child Node selection state(s)."

End Sub

Public Sub ToggleSelection(Optional ByRef Node As MSComctlLib.Node)
Attribute ToggleSelection.VB_Description = "Toggles the MultiSelect selection."

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

End Function
