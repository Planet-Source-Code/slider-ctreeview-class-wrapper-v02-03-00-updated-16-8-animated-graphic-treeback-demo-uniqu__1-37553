VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cTabStrip
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         01/08/2002
' Version:      01.00.00
' Description:  Hooks the TabStrip control to manage it's associated controls
'               and TabStrip tab clicking. Essentially removing the need of
'               control from the developer to an automatic class wrapper.
' Edit History: 01.00.00 01/08/2002 Initial Release
'
' Notes:        To use, declare & initialise the class, hook the TabStrip
'               control, add associated controls (for multiple controls for
'               a specific tab, put the controls on a container control
'               (like a PictureBox or Frame) and associate it) and
'               management of the tab selection process is automatic!
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 1                    '## 0=No debug
                                        '   1=debug
#Const MOUSEEVENTS = 1                  '## 0=No mouse events
                                        '   1=Mouse Up & Mouse Down
                                        '   2=All Mouse events
#If DEBUGMODE = 1 Then
    Private dbgCtrlName  As String
#End If

'===========================================================================
' Private: API Delclarations & variables
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                        (lpvDest As Any, _
                                         lpvSource As Any, _
                                         ByVal cbCopy As Long)

'===========================================================================
' Private: Variables and Declarations
'
Private WithEvents moTabStrip As MSComctlLib.TabStrip
Attribute moTabStrip.VB_VarHelpID = -1

Private mbAutoFit     As Boolean
Private meAlignment   As eAlignment
Private mlOldTabIndex As Long
Private moObjPtr()    As Long           '## Array of object pointers

'===========================================================================
' Private: cTabStrip Internal Error Declarations
'
Private Const csSOURCE_ERR   As String = "cTabStrip"
Private Const clTABCONT_ERR  As Long = vbObjectError + 2048 + 210
Private Const csTABCONT_ERR  As String = "Invalid Container control"
Private Const clTABNDX_ERR  As Long = vbObjectError + 2048 + 211
Private Const csTABNDX_ERR  As String = "Invalid Tab reference - "
'Private Const clTABHOOK_ERR  As Long = vbObjectError + 2048 + 203
'Private Const csTABHOOK_ERR  As String = "Invalid control! Must be a VB6 TabStrip control"

'===========================================================================
' Public : Enumerations
'
Public Enum eAlignment
    [Align TopLeft] = 1
    [Align TopHCenter] = 2
    [Align TopRight] = 3
    [Align LeftVCenter] = 4
    [Align VCenterHCenter] = 5
    [Align RightVCenter] = 6
    [Align BottomLeft] = 7
    [Align BottomHCenter] = 8
    [Align BottomRight] = 9
End Enum

'===========================================================================
' TabStrip: Events
'
Private Sub moTabStrip_BeforeClick(Cancel As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::BeforeClick -> Cancel="; CStr(Cancel)
    #End If
End Sub

Private Sub moTabStrip_Click()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Click"
    #End If
    '
    '## A Tab has been selected. Show the correct attached control and hide
    '   the old one.
    '
    pRefresh

End Sub

Private Sub moTabStrip_DragDrop(Source As Control, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragDrop -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If
End Sub

Private Sub moTabStrip_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragOver -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If
End Sub

Private Sub moTabStrip_GotFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::GotFocus"
    #End If
End Sub

Private Sub moTabStrip_KeyDown(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyDown -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If
End Sub

Private Sub moTabStrip_KeyPress(KeyAscii As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyPress -> KeyAscii="; CStr(KeyAscii)
    #End If
End Sub

Private Sub moTabStrip_KeyUp(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyUp -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If
End Sub

Private Sub moTabStrip_LostFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::LostFocus"
    #End If
End Sub

Private Sub moTabStrip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseDown -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
End Sub

Private Sub moTabStrip_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseMove -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
End Sub

Private Sub moTabStrip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseUp -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
End Sub

Private Sub moTabStrip_OLECompleteDrag(Effect As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLECompleteDrag -> Effect="; CStr(Effect)
    #End If
End Sub

Private Sub moTabStrip_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragDrop -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
    #End If
End Sub

Private Sub moTabStrip_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragOver -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y); "  State="; CStr(State)
    #End If
End Sub

Private Sub moTabStrip_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEGiveFeedback -> Effect="; CStr(Effect); "  DefaultCursors="; CStr(DefaultCursors)
    #End If
End Sub

Private Sub moTabStrip_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLESetData -> Effect="; CStr(DataFormat)
    #End If
End Sub

Private Sub moTabStrip_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEStartDrag -> AllowedEffects="; CStr(AllowedEffects)
    #End If
End Sub

Private Sub moTabStrip_Validate(Cancel As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Validate -> Cancel="; CStr(Cancel)
    #End If
End Sub

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    ReDim moObjPtr(0)   '## Initialise
    mbAutoFit = True
    meAlignment = [Align VCenterHCenter]
End Sub

Private Sub Class_Terminate()
    '## release control
    UnHookCtrl
End Sub

'===========================================================================
' Public Properties
'
Public Property Get Alignment() As eAlignment
Attribute Alignment.VB_Description = "Set/Get positioning of attached controls within the TabStrip's client area."
    Alignment = meAlignment
End Property

Public Property Let Alignment(ByVal NewState As eAlignment)
    meAlignment = NewState
    If Not mbAutoFit Then pRefresh
End Property

Public Property Get AutoFit() As Boolean
Attribute AutoFit.VB_Description = "Get/Set whether attached controls are automatically resized to the TabStrip's client area."
    AutoFit = mbAutoFit
End Property

Public Property Let AutoFit(ByVal NewState As Boolean)
    mbAutoFit = NewState
    If mbAutoFit Then pRefresh
End Property

Public Property Set TabStipContainer(oControl As Object)
Attribute TabStipContainer.VB_Description = "Sets a new container for the TabStrip and all attached controls."

    On Error GoTo ErrorHandler

    Dim iLoop As Long
    Dim lPtr  As Long
    '
    '## Set the TabStrip to the new container control and also for
    '   every associated control.
    '
    Set moTabStrip.Container = oControl
    For iLoop = 1 To UBound(moObjPtr)
        lPtr = moObjPtr(iLoop)
        '
        '## Only set containers if a valid control pointer exists
        '
        If lPtr Then
            Set pObj(lPtr).Container = oControl
        End If
    Next
    '
    '## Most likely (as the TabStrip has no _Resize event) the TabStrip was resized when
    '   its container  was changed.
    '
    pRefresh
    Exit Property

ErrorHandler:
    Err.Raise clTABCONT_ERR, csSOURCE_ERR, csTABCONT_ERR
End Property

'===========================================================================
' Public subroutines and functions
'
Public Sub Attach(ByVal TabItem As Variant, ByRef Ctrl As VB.Control)
Attribute Attach.VB_Description = "Attach a control to a Tab."

    Dim lTabNdx As Long

    On Error GoTo ErrorHandler

    If Not (Ctrl Is Nothing) Then
        '
        '## Re-align our internal object pointer array with the number of Tabs in the
        '   TabStrip control
        '
        pUpdateBounds
        '
        '## To aviod Error 13 (Type Mismatch) with invalid keys, we need to safely
        '   test if the key exists - if not, then raise the appropriate error message.
        '
        If VarType(TabItem) = vbString Then
            If Not pExist(TabItem) Then
                Err.Raise 35600, , "Invalid Key (key not found)"
            End If
        End If
        '
        '## Get the referenced Tab's index
        '
        lTabNdx = moTabStrip.Tabs(TabItem).Index
        If moObjPtr(lTabNdx) Then
            '
            '## Tab already has a control attached, therefore detach
            '   it before setting a new association.
            '
            Detach TabItem
        End If
        '
        '## Store a weak object reference to the control to aviod circular references
        '
        moObjPtr(lTabNdx) = ObjPtr(Ctrl)
        With Ctrl
            '
            '## Set the attached control's container to math that of the TabStrip
            '   to ensure they're the same.
            '
            Set .Container = moTabStrip.Container
            '
            '## Hide the attached control until we need it.
            '
            .Visible = False
        End With
        pRefresh
    End If
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 35600
            Err.Raise clTABNDX_ERR, csSOURCE_ERR, csTABNDX_ERR + Err.Description
        Case Else
            MsgBox Err.Number + " - " + Err.Description
    End Select
End Sub

Public Sub Detach(ByVal TabItem As Variant)
Attribute Detach.VB_Description = "Detach a control from a Tab."

    Dim lTabNdx As Long

    lTabNdx = moTabStrip.Tabs(TabItem).Index
    If pCheckUBound Then
        '
        '## It's safe to drop the control reference
        '
        If moObjPtr(lTabNdx) Then
            pObj(moObjPtr(lTabNdx)).Visible = False
            moObjPtr(lTabNdx) = 0
        End If
    End If

End Sub

Public Sub HookCtrl(ByRef Ctrl As MSComctlLib.TabStrip)
Attribute HookCtrl.VB_Description = "Primes the cTabStrip class to manage the Tabstrip control and attached controls."

    Set moTabStrip = Ctrl                           '## Capture the TabStrip control.
                                                    '   Class won't work if this isn't called first.
    With moTabStrip
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
        'mlHwnd = .hWnd
    End With
    mlOldTabIndex = moTabStrip.SelectedItem.Index
    pUpdateBounds

End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Called if the TabStrip is resized. (Note: The TabStrip has no Resize Event.)"
    pRefresh
End Sub

Public Sub UnHookCtrl()
Attribute UnHookCtrl.VB_Description = "Releases the hooked TabStripcontrol."
    '
    '## Release the hooked control
    '
    Set moTabStrip = Nothing
    '
    '## Re-initialise object pointer array - don't want to keep old pointers
    '
    ReDim moObjPtr(0)

End Sub

'===========================================================================
' General: Private subroutines and functions
'
Public Function pCheckUBound() As Boolean
    '
    '## Make sure that there will be no illegal Tab references
    '
    If moTabStrip.Tabs.Count >= UBound(moObjPtr) Then pCheckUBound = True

End Function

Private Function pExist(TabItem As Variant) As Boolean

    On Error Resume Next
    pExist = (moTabStrip.Tabs(TabItem).Index > 0)

End Function

Private Function pObj(ByVal lpObj As Long) As VB.Control
    '
    '## This function takes a dumb numeric pointer and turns it into
    '   a valid object reference.
    '
    '   WARNING: *** Do not end your project while in this sub unless you like to crash. ***
    '
    Dim oCtrl As VB.Control

    CopyMemory oCtrl, lpObj, 4&
    Set pObj = oCtrl
    CopyMemory oCtrl, 0&, 4&

End Function

Private Sub pRefresh()

    Dim oCtrl As VB.Control

    If mlOldTabIndex Then
        '
        '## Hide currently visible control (if one is referenced)
        '
        If moObjPtr(mlOldTabIndex) Then
            pObj(moObjPtr(mlOldTabIndex)).Visible = False
        End If
    End If
    '
    '## Get new Tab Index
    '
    mlOldTabIndex = moTabStrip.SelectedItem.Index
    '
    '## Is there a valid control reference?
    '
    If moObjPtr(mlOldTabIndex) Then
        '
        '## Reference the active attached control
        '
        Set oCtrl = pObj(moObjPtr(mlOldTabIndex))
        '
        '## Position the active attached control
        '
        With moTabStrip
            If mbAutoFit Then
                oCtrl.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
            Else
                If oCtrl.Height > .ClientHeight Then
                    oCtrl.Height = .ClientHeight
                End If
                If oCtrl.Width > .ClientWidth Then
                    oCtrl.Width = .ClientWidth
                End If
                Select Case meAlignment
                    Case [Align TopLeft]
                        oCtrl.Move .ClientLeft, _
                                   .ClientTop
                    Case [Align TopHCenter]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width) / 2, _
                                   .ClientTop
                    Case [Align TopRight]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width), _
                                   .ClientTop
                    Case [Align LeftVCenter]
                        oCtrl.Move .ClientLeft, _
                                   .ClientTop + (.ClientHeight - oCtrl.Height) / 2
                    Case [Align VCenterHCenter]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width) / 2, _
                                   .ClientTop + (.ClientHeight - oCtrl.Height) / 2
                    Case [Align RightVCenter]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width), _
                                   .ClientTop + (.ClientHeight - oCtrl.Height) / 2
                    Case [Align BottomLeft]
                        oCtrl.Move .ClientLeft, _
                                   .ClientTop + (.ClientHeight - oCtrl.Height)
                    Case [Align BottomHCenter]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width) / 2, _
                                   .ClientTop + (.ClientHeight - oCtrl.Height)
                    Case [Align BottomRight]
                        oCtrl.Move .ClientLeft + (.ClientWidth - oCtrl.Width), _
                                   .ClientTop + (.ClientHeight - oCtrl.Height)
                End Select
            End If
        End With
        '
        '## Bring to front & show the active attached control
        '
        With oCtrl
            .ZOrder
            .Visible = True
            '
            '## if the active attached control is a TabStrip, then tell it to do
            '   it's housekeeping by forcing the _Click event as the TabStrip control
            '   has no _Paint event.
            '
            If TypeOf oCtrl Is MSComctlLib.TabStrip Then
                .SelectedItem = .SelectedItem
            End If
        End With
    End If

End Sub

Private Sub pUpdateBounds()
    '
    '## Re-align our internal object pointer array with the number of Tabs in the
    '   TabStrip control
    '
    Dim lCount As Long

    lCount = moTabStrip.Tabs.Count
    If lCount <> UBound(moObjPtr) Then
        ReDim Preserve moObjPtr(lCount)
    End If

End Sub
