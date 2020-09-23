VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:  cTags
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        03/02/2001
' Version:     01.00.00
' Description: Variant Collection class
' Notes:       Collection class offered more functionality over an array of
'              UDTs (User Defined Type variable structure) or an array of
'              variants.
'              -=-=-= Warning! =-=-=-
'              This class requires 'Break on Unhandled Errors' if run in the
'              IDE. To set this, on the IDE menu:- Tools > Options > set
'              Error Trapping = Break on Unhandled Errors option.
'
'===========================================================================

Option Explicit

Private moTagCol As Collection

Public Sub Add(ByVal lPtr As Long, ParamArray Elements() As Variant)
    If UBound(Elements) > 0 Then
        If Not Exist(lPtr) Then
            moTagCol.Add Elements, CStr(lPtr)   '## CStr(lPtr) is used as straight numeric
                                                '   values can't be used as a collection key
            'Debug.Print "TAG::ADD>> " + CStr(Elements(0)) + ", " + CStr(Elements(1)) + ", " + _
                                         CStr(Elements(2)) + ", " + CStr(Elements(3))
        End If
    End If
End Sub

Public Sub Remove(ByVal lPtr As Long)
    If Exist(lPtr) Then
        moTagCol.Remove CStr(lPtr)
    End If
End Sub

Public Function Exist(ByVal lPtr As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim Tmp As Variant
    Tmp = moTagCol(CStr(lPtr))(0)               '## (0) = variant array element zero
    Exist = True
ErrorHandler:
End Function

Public Property Get Count() As Long
    Count = moTagCol.Count
End Property

Public Property Get Element(ByVal lPtr As Long, ByVal ElementID As Long) As Variant
Attribute Element.VB_UserMemId = 0
    Element = moTagCol(CStr(lPtr))(ElementID)   '## (ElementID) = variant array element ElementID
End Property

Public Property Let Element(ByVal lPtr As Long, ByVal ElementID As Long, NewData As Variant)
    moTagCol(CStr(lPtr))(ElementID) = NewData
End Property

Public Property Set Element(ByVal lPtr As Long, ByVal ElementID As Long, NewObj As Variant)
    Set moTagCol(CStr(lPtr))(ElementID) = NewObj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = moTagCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set moTagCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set moTagCol = Nothing
End Sub
