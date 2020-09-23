VERSION 5.00
Begin VB.Form fViewBMP 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "View Image: "
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2115
   Icon            =   "fViewBMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox picDialog 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   105
      ScaleHeight     =   1005
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   105
      Width           =   1515
   End
End
Attribute VB_Name = "fViewBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    frmViewBitmap.frm
' Author:       Graeme Grant
' Date:         1/5/2001
' Version:      01.00.00  (TKWIN v1.10)
' Description:  Main Wrapper for viewing bitmaps. Does not validate the
'               bitmap colour depth or dimensions.
' Edit History:
'
'===========================================================================

Option Explicit

Private Type Rect
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Const SPI_GETWORKAREA As Long = 48                          '@@ v01.01.00

Private Sub cmdDialog_Click()
    Unload Me
End Sub

Public Property Set PicBox(New_Bitmap As VB.PictureBox)

    Dim lGap    As Long
    Dim lwidth  As Long
    Dim lheight As Long

    On Error Resume Next

    lGap = Screen.TwipsPerPixelX * 5
    '
    '## Reposition the bitmap on the form
    '
    With picDialog
        Set .Picture = New_Bitmap.Picture
        .Move lGap, lGap, .Width, .Height
        lwidth = .Width
        lheight = .Height
    End With
    '
    '## Reposition button so that it's not over the image
    '
    With cmdDialog
        .Move lGap + (lwidth - .Width) \ 2, picDialog.Top + lheight + lGap * 3, .Width, .Height
    End With
    '
    '## Resize form to fit bitmap
    '
    With Me
        .Move .Left, .Top, lwidth + lGap * 3.5, lheight + cmdDialog.Height + lGap * 5 + 500
    End With
    '
    '## Position the form on the screen
    '
    CenterForm

End Property

Private Sub CenterForm()

    Dim tCRect   As Rect        '## Holds the area that the form is to be centered in
    Dim x        As Single
    Dim y        As Single

    SystemParametersInfo SPI_GETWORKAREA, 0&, VarPtr(tCRect), 0&    '@@ v01.01.00
    '
    '## Center the Form
    '
    With Me
        x = (((tCRect.Right - tCRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
        y = (((tCRect.Bottom - tCRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
        .Move x, y
    End With

End Sub
