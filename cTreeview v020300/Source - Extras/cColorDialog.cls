VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cColorDialog
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         23/03/2002
' Version:      01.00.00
' Description:  Wrapper class for Microsoft's Color Picker Common Dialog.
' Edit History: 01.00.00 23/03/2002 Initial Release
'
'===========================================================================

Option Explicit

'## ChooseColor structure flag constants
Public Enum eColorFlags
    CC_RGBINIT = &H1
    CC_FULLOPEN = &H2
    CC_PREVENTFULLOPEN = &H4
    CC_SOLIDCOLOR = &H80
    CC_ANYCOLOR = &H100
End Enum

'## Static array to contain the custom colours selected by the user
Private mlCustClrs(0 To 15) As OLE_COLOR
Private meColorFlags        As eColorFlags
Private mlInitColor         As OLE_COLOR

Private Type CHOOSECOLORSTRUCT
   lStructSize     As Long
   hwndOwner       As Long
   hInstance       As Long
   rgbResult       As Long
   lpCustColors    As Long
   flags           As Long
   lCustData       As Long
   lpfnHook        As Long
   lpTemplateName  As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" _
                                   Alias "ChooseColorA" (lpcc As CHOOSECOLORSTRUCT) As Long


Public Function Execute(Optional Host As Object) As OLE_COLOR

    Dim cc As CHOOSECOLORSTRUCT

    With cc
        '## Set flag values and Initial color
        .flags = CC_ANYCOLOR Or meColorFlags
        If (meColorFlags And CC_RGBINIT) = CC_RGBINIT Then
            .rgbResult = mlInitColor
        End If

        '## Size of structure
        .lStructSize = Len(cc)

        '## Owner of the dialog
        If Not (Host Is Nothing) Then
            .hwndOwner = Host.hWnd
        End If

        '## Assign the custom colour selections
        .lpCustColors = VarPtr(mlCustClrs(0))

    End With

    If ChooseColor(cc) = 1 Then
        Execute = cc.rgbResult
    End If

End Function

Public Property Get ColorFlags() As eColorFlags
    ColorFlags = meColorFlags
End Property

Public Property Let ColorFlags(ByVal NewFlags As eColorFlags)
    meColorFlags = NewFlags
End Property

Public Property Get CustomColor(ByVal Index As Long) As OLE_COLOR
    CustomColor = mlCustClrs(Index)
End Property

Public Property Let CustomColor(ByVal Index As Long, ByVal vNewValue As OLE_COLOR)
    mlCustClrs(Index) = vNewValue
End Property

Public Property Get InitColor() As OLE_COLOR
    InitColor = mlInitColor
End Property

Public Property Let InitColor(ByVal NewColor As OLE_COLOR)
    mlInitColor = NewColor
End Property

Private Sub Class_Initialize()

  '## Initialize the custom colours with a series of gray shades
   Dim lLoop As Long
   For lLoop = 240 To 15 Step -15
      mlCustClrs((lLoop \ 15) - 1) = RGB(lLoop, lLoop, lLoop)
   Next

End Sub
