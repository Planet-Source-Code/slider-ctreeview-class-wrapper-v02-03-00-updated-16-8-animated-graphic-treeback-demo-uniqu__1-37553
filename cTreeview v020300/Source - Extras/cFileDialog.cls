VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'------------------------------------------------------------------
'   Name:           CFileDialog (CFileDialog.CLS)
'   Type:           Utility class wrapper
'   Description:    Procedures for accessing the File Common Dialog services through the Windows API
'                   without the overhead of the VB Common Dialog OCX control.
'
'   Author:         Klaus H. Probst [kprobst@vbbox.com]
'   URL:            http://www.vbbox.com/
'   Copyright:      This work is copyright © 1998-2001, Klaus H. Probst
'   Usage:          You may use this code as you see fit, provided that you assume all
'                   responsibilities for doing so.
'   Distribution:   If you intend to distribute the file(s) that make up this sample to
'                   any WWW site, online service, electronic bulletin board system (BBS),
'                   CD or any other electronic or physical media, you must notify me in
'                   advance to obtain my express permission.
'
'   Notes:
'
'       The .FileName property is the standard entry point to this class.
'       After you've initialized all relevant property members, just assign
'       the value of the FileName property to a variable. The class will open
'       the dialog and return the selected file path, or an empty string if the
'       user canceled or there was an error.
'       The following illustrates how you might use the class:
'
'        Dim hDlg As CFileDialog
'        Dim strFileName As String
'
'        Set hDlg = New CFileDialog
'
'        With hDlg
'
'            .DialogMode = cdlgOpen
'            .Path = "C:\My Documents"
'            .Flags = cdlgOFNFileMustExist Or cdlgOFNHideReadOnly Or cdlgOFNExplorer
'            .hwndOwner = Me.hWnd
'            .Mask = "Word Documents (*.doc)|*.doc|All Files (*.*)|*.*"
'            .Title = "Select Document"
'            strFileName = .FileName
'
'        End With
'
'        Set hDlg = Nothing
'
'       If you include the OFN_ALLOWMULTISELECT bit in the dialog flags, make sure
'       you use a variant instead of a string when getting the value of FileName,
'       because that's what you will get back from the .FileName property (a variant
'       array). You can then iterate through each index in the array using UBound to
'       retreive each filename.
'       Please see the code comments and method headers for specific info on how the
'       class works.
'
'
'   Dependencies:
'
'       (none)
'
'------------------------------------------------------------------------------------------------------
Option Explicit
DefLng A-Z

#If Not CB_NO_LOCAL_SYMBOLS Then

Private Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'//GetOpen and GetSaveFileName() constants
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

'// Declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As tagOPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As tagOPENFILENAME) As Long
Private Declare Function CommDlg_GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Long) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

#End If

'// Module-wide
Private m_ofn As tagOPENFILENAME


'// Default "All Files" filter mask
Private Const m_strcdlgDefMask = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar

'// possible operation modes: "Open" or "Save"
Public Enum FileDialogModes
    cdlgOpen = 1
    cdlgSave = 2
End Enum

Private m_Mode As FileDialogModes


Private Const OFN_SAVEDEFAULT = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or _
                                OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_NOCHANGEDIR Or _
                                OFN_CREATEPROMPT
    
Private Const OFN_OPENDEFAULT = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or _
                                OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_NOCHANGEDIR


Public Enum FileDialogFlags

    cdlgOFNReadOnly = OFN_READONLY
    cdlgOFNOverwritePrompt = OFN_OVERWRITEPROMPT
    cdlgOFNHideReadOnly = OFN_HIDEREADONLY
    cdlgOFNNoChangeDir = OFN_NOCHANGEDIR
    cdlgOFNShowHelp = OFN_SHOWHELP
    cdlgOFNNoValidate = OFN_NOVALIDATE
    cdlgOFNAllowMultiSelect = OFN_ALLOWMULTISELECT
    cdlgOFNPathMustExist = OFN_PATHMUSTEXIST
    cdlgOFNFileMustExist = OFN_FILEMUSTEXIST
    cdlgOFNCreatePrompt = OFN_CREATEPROMPT
    cdlgOFNShareAware = OFN_SHAREAWARE
    cdlgOFNNoReadOnlyReturn = OFN_NOREADONLYRETURN
    cdlgOFNNoTestFileCreate = OFN_NOTESTFILECREATE
    cdlgOFNNoNetworkButton = OFN_NONETWORKBUTTON
    cdlgOFNNoLongNames = OFN_NOLONGNAMES
    cdlgOFNExplorer = OFN_EXPLORER
    cdlgOFNNoDereferenceLinks = OFN_NODEREFERENCELINKS
    cdlgOFNLongNames = OFN_LONGNAMES
    
    cdlgOFNSaveDefault = OFN_SAVEDEFAULT
    cdlgofnOpenDefault = OFN_OPENDEFAULT
    
End Enum

'// CommDlg errors returned by CommDlgExtendedError()
'// commented values correspond to hook and template
'// errors, which we don't use at all. Note that the
'// CDERR_CANCELED value is not defined by the Win32 API,
'// I added it to be able to discern between an error and
'// a canceled dialog.
Public Enum CommonDialogFileErrors

    CDERR_CANCELED = 0&

    CDERR_DIALOGFAILURE = &HFFFF
    
    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    'CDERR_NOTEMPLATE = &H3
    'CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    'CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC
    
    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003
    
End Enum

'// if the call to GetxxxxFileName() fails, we save this here
Private m_enlastError As CommonDialogFileErrors

'
'   Read only. If the call to GetxxxxFileName() fails
'   when reading the FileName property, the class will call
'   CommDlgExtendedError() and store the error value, which
'   you can retreive later.
'
'
Public Property Get LastError() As CommonDialogFileErrors

    LastError = m_enlastError

End Property


'
'   Read only. Returns the zero-based position
'   in the FileName string where the path ends
'   and the actual filename begins.
'
Public Property Get FileOffset() As Long

    FileOffset = m_ofn.nFileOffset

End Property

'
'   Read only. Returns the zero-based position
'   in the FileName string where the filename ends
'   and the extension begins.
'
Public Property Get ExtensionOffset() As Long

    ExtensionOffset = m_ofn.nFileExtension

End Property

'
'   Read only. Returns any custom filters the user
'   may have entered while in the dialog. Note that
'   even if the user does not enter any different filters
'   than those on the Mask property, this will be filled
'   with the last filter selected. In other words, if the
'   first filter in the Mask property is "*.DOC", and the
'   user just selects a file, this property will be "*.doc".
'   Also, note that user-entered filters are modified by
'   the dialog in that the extension will have a wildcard
'   appended if the user does not enter one. This means that
'   if the user entered "*.xls", the dialog will return
'   "*.xls*". This is even if the wildcard is obviously
'   not desired or needed by the user. Weird, but true.
'
Public Property Get CustomFilter() As String

    CustomFilter = m_ofn.lpstrCustomFilter

End Property


'
'   Read only. After showing the dialog, this
'   will contain the display title of the selected
'   file (the string that shows up on the Explorer
'   views, depending on whether or not registered
'   extensions are displayed by the shell, and so on),
'   except if there was a multiple selection, in which
'   case it will be empty. For this cases, you can
'   use the GetFileTitle() function.
'
Public Property Get FileTitle() As String

    FileTitle = m_ofn.lpstrFileTitle

End Property

'
'   Sets/returns the flags that make the dialog tick.
'   As with the OCX, you must OR the enumerated constants
'   together. if you don't assign anything here, the class
'   will use its defaults, which are defined in the [Declarations]
'   section
'
Public Property Get Flags() As FileDialogFlags

    Flags = m_ofn.Flags

End Property

'
'
'
Public Property Let Flags(ByVal rValue As FileDialogFlags)

    m_ofn.Flags = rValue

End Property



'
'   Sets/returns the mode of the dialog, either
'   "Open" or "Save".
'
Public Property Get DialogMode() As FileDialogModes

    DialogMode = m_Mode

End Property

'
'
'
Public Property Let DialogMode(ByVal rValue As FileDialogModes)

    m_Mode = rValue

    '// this resets the flags, so make sure you set this
    '// before the flags or reset the flags after this.
    If m_Mode = cdlgOpen Then
        m_ofn.Flags = cdlgofnOpenDefault
    
    ElseIf m_Mode = cdlgSave Then
        m_ofn.Flags = cdlgOFNSaveDefault
    
    Else
        m_ofn.Flags = 0

    End If

End Property

'
'   Sets/returns the index (zero-based) of the filter
'   contained in the Mask property that will be initially
'   used by the dialog. Default is, of course, zero.
'
Public Property Get DefaultFilterIndex() As Integer

    DefaultFilterIndex = m_ofn.nFilterIndex

End Property
'
'
'
Public Property Let DefaultFilterIndex(ByVal rValue As Integer)

    m_ofn.nFilterIndex = rValue

End Property


'
'   Sets/returns the default extension for the
'   file(s) selected in the dialog. You can leave
'   this alone if you don't need it.
'
Public Property Get Extension() As String

    Extension = m_ofn.lpstrDefExt

End Property
'
'
'
Public Property Let Extension(ByVal rValue As String)

    m_ofn.lpstrDefExt = rValue

End Property


'
'   Sets/returns the form or window that will act
'   as "parent" of the dialog so that it is shown
'   modally. If you leave this alone, the class will
'   attempt to use Screen.ActiveForm to obtain a valid
'   hWnd. If you don't want this, assign a -1 to this
'   property to force a non-modal display.
'
Public Property Get hwndOwner() As Long

    hwndOwner = m_ofn.hwndOwner

End Property
'
'
'
Public Property Let hwndOwner(ByVal rValue As Long)

    m_ofn.hwndOwner = rValue

End Property


'
'   Sets/returns the title of the dialog box. Default is "Open"
'   for cdlgOpen and "Save" for cdlgSave modes. Not to be confused
'   with the FileTitle property.
'
Public Property Let Title(ByVal rValue As String)

    m_ofn.lpstrTitle = rValue

End Property

'
'
'
Public Property Get Title() As String

    Title = m_ofn.lpstrTitle

End Property


'
'   You can set this to initialize the
'   dialog. If it is empty, the value of CurDir$
'   is used (by the CommDlg DLL, not the class)
'
Public Property Get Path() As String

    Path = m_ofn.lpstrInitialDir

End Property

'
'
'
Public Property Let Path(ByVal rValue As String)

    m_ofn.lpstrInitialDir = rValue

End Property


'
'   You can set this to a valid filename to initialize
'   the dialog.
'
Public Property Let FileName(ByVal rValue As Variant)

    If VarType(rValue) = vbString Then m_ofn.lpstrFile = rValue

End Property

'
'   Standard entry point for the class. Upon
'   using this property in an assignment, the class
'   shows the dialog and returns whatever the user
'   selected, or an emtpty string is it was canceled
'   or an error ocurred. Note that if you're using
'   multiple selections, you must do the assignment
'   to a variant (instead of a string) for this to work.
'
Public Property Get FileName() As Variant

    If GetOpenSaveFileName() = True Then

        If (m_ofn.Flags And OFN_ALLOWMULTISELECT) Then
            FileName = ResolveMultiSelection()
    
        Else
            FileName = m_ofn.lpstrFile
    
        End If

    Else
        FileName = ""

    End If

End Property



'
'   Sets/returns the mask or filter mask that the dialog
'   will use. A valid mask string looks like this:
'
'       Databases (*.mdb;*.mda;*.mdw)|*.mdb;*.mda;*.mdw|All files (*.*)|*.*
'
'   Note the use of pipe characters ("|", Chr$(124)) as delimiters
'   between filters. If this is an empty string, it is set to "*.*"
'
Public Property Get Mask() As String

    Mask = m_ofn.lpstrFilter

End Property

'
'
'
Public Property Let Mask(ByVal rValue As String)

    m_ofn.lpstrFilter = rValue

End Property


'
'   Resolves individual char delimited filenames from
'   a multiple-selection GetOpenFileName call and
'   returns a variant array with each element as
'   an array index
'
Private Function ResolveMultiSelection() As Variant

    Dim arrFileNames() As String
    Dim intNullPos As Integer
    Dim strSelection As String
    Dim intCounter As Integer
    Dim charDelimiter As String
    Dim strPath As String

    'Debug.Assert 0

    On Error GoTo catch

    strSelection = m_ofn.lpstrFile
    If Len(strSelection) = 0 Then Exit Function

    ReDim arrFileNames(0 To 0)

    '// If the user selected more than one file, the string consists of the
    '// path followed by the single null-separated names of all the files.
    '// In the above case, the nFileOffset member of OPENFILENAME contains
    '// the position of the character where the filenames start.
    '// If the user selected only one file, then nFileExtension is non zero and the
    '// lpstrFile member contains the normal-style Path+File format followed by
    '// the usual NULLs.

    '// So, as per above, look at the nFileExtension member. If it is non-zero, we just
    '// get the one complete filename, trim the nulls and assign it to the array (since
    '// we need to return an array in any case).
    '// Otherwise continue normally.
    If m_ofn.nFileExtension <> 0 Then
        intNullPos = InStr(strSelection, vbNullChar)
        arrFileNames(0) = VBA.Left$(strSelection, intNullPos - 1)
    
    Else

        '// If we're working with the Explorer-type dialog,
        '// the delimiter between elements is a NULL char.
        '// Otherwise it's a space, since old-style dialogs
        '// don't return long filenames.
        If (m_ofn.Flags And cdlgOFNExplorer) Then
            charDelimiter = vbNullChar
        Else
            charDelimiter = Chr$(32)
        End If
    
        '// The first null delimiter is the path, so get it.
        intNullPos = m_ofn.nFileOffset
    
        '// Save this
        m_ofn.lpstrInitialDir = VBA.Left$(strSelection, intNullPos - 1)
    
        strPath = m_ofn.lpstrInitialDir
    
        If VBA.Right$(strPath, 1) <> "\" Then _
           strPath = strPath & "\"
    
        '// re-trim
        strSelection = Mid$(strSelection, intNullPos + 1)
    
        Do
    
            intNullPos = InStr(strSelection, charDelimiter)
        
            '// In non-explorer mode, the last filename
            '// is delimited by NULLs, *not* spaces
            If intNullPos = 0 And Not (m_ofn.Flags And cdlgOFNExplorer) Then intNullPos = InStr(strSelection, vbNullChar)
            If intNullPos <= 1 Then Exit Do         '// egads, we trimmed everything and there's only
                                                    '// NULLs left. This is less expensive than trimming
                                                    '// nulls off the string from the right.
        
            ReDim Preserve arrFileNames(0 To intCounter)
        
            '// append the path to each filename
            arrFileNames(intCounter) = strPath & VBA.Left$(strSelection, intNullPos - 1)
            strSelection = Mid$(strSelection, intNullPos + 1)
        
            intCounter = intCounter + 1
        
        Loop While Len(strSelection) > 2        '// because theoretically, the string is delimited
                                                '// by two ending NULLs or spaces
    End If    '// valid multiple selection

    '// return the array
    ResolveMultiSelection = arrFileNames()

cleanup:

    On Error Resume Next

    Erase arrFileNames

    Exit Function

catch:

    '// whooops
    Debug.Assert 0
    ResolveMultiSelection = Empty
    Resume cleanup

End Function

'
'   Creates a null-delimited filter mask string that the
'   GetxxxxFilename() APIs can understand from a pipe-delimited
'   "|" string that is more what the VB user is used to
'
Private Function MAKEFILTERMASK(ByVal szFilter As String) As String

    On Error GoTo catch

    Dim lpszOutMask As String
    Dim a&
    Dim intPipePos As Integer

    '// if the string is empty or null, return NULL;
    If szFilter = "" Then MAKEFILTERMASK = vbNullString: Exit Function

    '// Sample string:
    '// Databases (*.mdb;*.mda;*.mdw)|*.mdb;*.mda;*.mdw|All files (*.*)|*.*

    '// assign a temp mask string, since the param is passed by value
    lpszOutMask = Trim$(szFilter)

    For a& = 0 To Len(lpszOutMask)

        '// Find the next pipe character
        intPipePos = InStr(lpszOutMask, Chr$(124))
        
        If intPipePos <> 0 Then                 '// if there actually is a pipe left...
            
            Mid$(lpszOutMask, intPipePos, 1) = vbNullChar      '// replace it with a NULLL
        
        Else                                '// we're out of pipes
        
            '// append the last two delimiters and quit the loop
            lpszOutMask = lpszOutMask & vbNullChar & vbNullChar
            Exit For
        
        End If

    Next a&

    '// return the ANSI string
    MAKEFILTERMASK = lpszOutMask

    Exit Function

catch:
    
    '// return the default mask
    MAKEFILTERMASK = m_strcdlgDefMask

End Function


'
'   This is what actually does the work.
'
Private Function GetOpenSaveFileName() As Boolean

    'Debug.Assert 0

    Dim lngReturn As Long
    Dim lngBuffSize As Long
    Dim lptmpOFN As tagOPENFILENAME

    On Error Resume Next

    With m_ofn

        '// initialize these
        .lpstrFileTitle = String$(260, vbNullChar)
        .nMaxFileTitle = Len(.lpstrFileTitle)
    
        .lpstrCustomFilter = String$(60, vbNullChar)      '// this is arbitrary, really.
        .nMaxCustFilter = Len(.lpstrCustomFilter)
    
        '// If the user provided a filename for initialization, we pad the
        '// string with null chars to the size of the buffer so that whatever
        '// comes back will fit on it.
    
        '// IMPORTANT:  If you're using multiple selection and the call is failing,
        '//             try increasing the size of the buffer below from 1024 to
        '//             something like 2048.
        If (m_ofn.Flags And OFN_ALLOWMULTISELECT) Then
            lngBuffSize = 1024
        Else
            lngBuffSize = 260
        End If
    
        .lpstrFile = .lpstrFile & String$(lngBuffSize - Len(.lpstrFile), vbNullChar)
        .nMaxFile = Len(.lpstrFile)

        '// if the filter is NULL, assign the default "All Files"
        '// else, build the string from the incoming filter
        If Len(.lpstrFilter) = 0 Then
            .lpstrFilter = m_strcdlgDefMask
        Else
            .lpstrFilter = MAKEFILTERMASK(.lpstrFilter)
        End If

        '// if the flags come in as NULL, assign the default flags depending on the action
        If .Flags = 0 Then
            If m_Mode = cdlgOpen Then
                .Flags = cdlgofnOpenDefault
        
            ElseIf m_Mode = cdlgSave Then
                .Flags = cdlgOFNSaveDefault
        
            End If
    
        End If

        .lStructSize = Len(m_ofn)
    
        '// reset the error value
        m_enlastError = CDERR_CANCELED
    
        '// actually call the function and show the dialog box, depending
        '// on what the user set the DialogMode property to
        If m_Mode = cdlgOpen Then
            lngReturn = GetOpenFileName(m_ofn)
    
        ElseIf m_Mode = cdlgSave Then
            lngReturn = GetSaveFileName(m_ofn)
    
        End If

        '// if the function returned True
        If lngReturn = 1 Then
    
            '// return the stripped FileName. If the Allow Multiselect flag is set,
            '// return the whole thing as is, but clear the title property
            If (m_ofn.Flags And OFN_ALLOWMULTISELECT) Then
                .lpstrFileTitle = ""        '// this is... um, invalid? I never checked it
        
            Else
        
                '// return the full path and filename
                .lpstrFile = VBA.Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
            
                '// return the path only. Note that this will include the "\" at the end.
                .lpstrInitialDir = Mid$(.lpstrFile, 1, .nFileOffset)
            
                '// return the file title
                .lpstrFileTitle = VBA.Left$(.lpstrFileTitle, _
                   InStr(.lpstrFileTitle, vbNullChar) - 1)
            
                '// return the selected file's extension using the nFileExtension
                .lpstrDefExt = Mid$(VBA.Left$(.lpstrFile, _
                   InStr(.lpstrFile, vbNullChar) - 1), .nFileExtension + 1)
            
            End If
        
       
            '// return any custom filter the user may have entered. This string
            '// is a bit weird; it has a null at the beggining. Thus the "2" offset
            '// in the Mid$() and InStr() calls.
            .lpstrCustomFilter = Mid$(.lpstrCustomFilter, 2, InStr(2, .lpstrCustomFilter, vbNullChar) - 2)
        
        
            GetOpenSaveFileName = True
        
        Else
        
            '// EEEKKK! Get the error code and clean everything up
            m_enlastError = CommDlgExtendedError()
            LSet m_ofn = lptmpOFN
        
        End If


    End With        '// m_ofn

End Function

'
'   Returns the shell display string used for a given file
'   (you must pass a full path), which varies depending on
'   the user's preferences (for example, if registered extensions
'   are not displayed, etc.). Handy if you need to get titles
'   after a multiple-selection dialog was shown.
'
Public Function GetFileTitle(ByVal FileName As String) As String

    Dim szBuffer As String
    Dim lReturn As Long

    szBuffer = String$(255, vbNullChar)

    lReturn = CommDlg_GetFileTitle(FileName, szBuffer, Len(szBuffer))

    '// NOTE:   The SDK states that GetFileTitle() returns 0 if it succeeds,
    '//         a negative number if it fails, or a positive number if the
    '//         buffer is too small. But for some strange reason, this is
    '//         not the case at all. It suceeds, but it returns a ridiculous
    '//         large value that means... I don't know what it means. So just
    '//         make sure you check your return values from this.

    On Error Resume Next    '// the InStr() could fail, so...

    GetFileTitle = VBA.Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)


End Function
