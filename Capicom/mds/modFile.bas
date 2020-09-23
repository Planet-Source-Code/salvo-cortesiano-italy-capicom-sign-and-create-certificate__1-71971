Attribute VB_Name = "modFile"
Option Explicit

Private mstrCurrentFolder As String
Public Const MAX_PATH = 260

Public Enum FileCommands
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
    OFS_MAXPATHNAME = 128
End Enum

Public Type OPENFILENAME
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Enum FileAttributes
   FILE_ATTRIBUTE_ARCHIVE = &H20
   FILE_ATTRIBUTE_COMPRESSED = &H800
   FILE_ATTRIBUTE_HIDDEN = &H2
   FILE_ATTRIBUTE_NORMAL = &H80
   FILE_ATTRIBUTE_READONLY = &H1
   FILE_ATTRIBUTE_SYSTEM = &H4
End Enum

Public Enum FileFlags
   FILE_FLAG_WRITE_THROUGH = &H80000000
   FILE_FLAG_NO_BUFFERING = &H20000000
   FILE_FLAG_OVERLAPPED = &H40000000
   FILE_FLAG_RANDOM_ACCESS = &H10000000
   FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
   FILE_FLAG_DELETE_ON_CLOSE = &H4000000
End Enum

Public Enum CreationDisposition
   CREATE_NEW = 1
   CREATE_ALWAYS = 2
   OPEN_EXISTING = 3
   OPEN_ALWAYS = 4
   TRUNCATE_EXISTING = 5
End Enum

Public Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type

Public Enum FileSharing
   FILE_SHARE_READ = &H1
   FILE_SHARE_WRITE = &H2
End Enum

Public Enum DesiredAccess
   GENERIC_WRITE = &H40000000
   GENERIC_READ = &H80000000
End Enum

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' .... BrowserForFolders
Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_BROWSEINCLUDEURLS = 128
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4
Private Const BIF_VALIDATE = 32
'Public Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILEDA = 3
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Dim m_StartFolder As String
Dim bValidateFailed As Boolean
Public Function GetDestinationFile(ParentHWnd As Long) As String

On Error GoTo Error_GetDestinationFile

Dim lngRet As Long
Dim udtFile As modFile.OPENFILENAME

With udtFile
    .lStructSize = Len(udtFile)
    .hwndOwner = ParentHWnd
    .hInstance = App.hInstance
    .lpstrFilter = "All Files" & VBA.vbNullChar & "*.*" & VBA.vbNullChar
    .nFilterIndex = 1
    .lpstrFile = String$(1025, 0)
    '  Just to be safe...
    .nMaxFile = Len(.lpstrFile) - 1
    '  Add the title.
    .lpstrTitle = "Select Destination File..."
    '  Add the default extension.
   .lpstrDefExt = vbNullChar
   .lpstrInitialDir = mstrCurrentFolder
    '  Now add the flags.
    '  The path + file does not have to exist.
    '  We'll keep the current directory as-is.
    '  We may alter the contents of the file,
    '  so warn the user of an overwrite.
    .flags = .flags Or FileCommands.OFN_NOCHANGEDIR _
                    Or FileCommands.OFN_HIDEREADONLY _
                    Or FileCommands.OFN_OVERWRITEPROMPT
    lngRet = GetSaveFileName(udtFile)
End With

If lngRet > 0 Then
    '  We have a selected file.
    '  Note that lpstrFile is the full file name with path
    '  and some null characters.
    '  So we have to do some parsing...
    GetDestinationFile = Trim$(VBA.Left$(udtFile.lpstrFile, _
                               InStr(1, udtFile.lpstrFile, VBA.vbNullChar) - 1))
    SetCurrentDirectory GetDestinationFile
End If

Exit Function

Error_GetDestinationFile:

End Function

Public Function GetSourceFile(ParentHWnd As Long) As String

On Error GoTo Error_GetSourceFile

Dim lngRet As Long
Dim udtFile As modFile.OPENFILENAME

With udtFile
    .lStructSize = Len(udtFile)
    .hwndOwner = ParentHWnd
    .hInstance = App.hInstance
    .lpstrFilter = "All Files" & VBA.vbNullChar & "*.*" & VBA.vbNullChar
    .nFilterIndex = 1
    .lpstrFile = String$(1025, 0)
    '  Just to be safe...
    .nMaxFile = Len(.lpstrFile) - 1
    '  Add the title.
    .lpstrTitle = "Select Source File..."
    '  Add the default extension.
   .lpstrDefExt = vbNullChar
   .lpstrInitialDir = mstrCurrentFolder
    '  Now add the flags.
    '  The path + file must exist.
    '  We'll keep the current directory as-is.
    '  We may alter the contents of the file.
    .flags = .flags Or FileCommands.OFN_FILEMUSTEXIST _
                    Or FileCommands.OFN_PATHMUSTEXIST _
                    Or FileCommands.OFN_NOCHANGEDIR _
                    Or FileCommands.OFN_HIDEREADONLY
    lngRet = GetOpenFileName(udtFile)
End With

If lngRet > 0 Then
    '  We have a selected file.
    '  Note that lpstrFile is the full file name with path
    '  and some null characters.
    '  So we have to do some parsing...
    GetSourceFile = Trim$(VBA.Left$(udtFile.lpstrFile, _
                          InStr(1, udtFile.lpstrFile, VBA.vbNullChar) - 1))
    SetCurrentDirectory GetSourceFile
End If

Exit Function

Error_GetSourceFile:

End Function

Private Sub SetCurrentDirectory(ReturnedFile As String)

On Error GoTo Error_SetCurrentDirectory

Dim strDir() As String
Dim X As Long

mstrCurrentFolder = vbNullString

strDir = Split(ReturnedFile, "\")

For X = LBound(strDir) To (UBound(strDir) - 1)
    mstrCurrentFolder = mstrCurrentFolder & strDir(X) & "\"
Next X

Exit Sub

Error_SetCurrentDirectory:

End Sub


Public Function BrowseFolder(ByVal strTitle As String, Optional strPath As String = "") As String
    Dim fOlder As String
    On Local Error GoTo ErrorHandler
    If strPath = "" Then
        strPath = App.Path + "\"
    Else
        If Right$(strPath, 1) <> "\" Then strPath = strPath + "\"
    End If
    fOlder = BrowseForFolder(frmMain.hwnd, strTitle, strPath)
    If fOlder <> "" Then BrowseFolder = fOlder Else BrowseFolder = ""
Exit Function
ErrorHandler:
        BrowseFolder = "Error!"
    Err.Clear
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Error Resume Next
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
    Select Case uMsg
        Case BFFM_INITIALIZED
            SendMessageA hwnd, BFFM_SETSELECTION, 1, m_StartFolder
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                SendMessageA hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer
            End If
        Case BFFM_VALIDATEFAILEDA
            bValidateFailed = True
    End Select
    BrowseCallbackProc = 0
End Function

Private Function BrowseForFolder(ByVal hwndOwner As Long, ByVal Prompt As String, Optional ByVal StartFolder) As String
    Dim lNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    On Local Error Resume Next
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = Prompt
        .ulFlags = BIF_BROWSEINCLUDEURLS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
        If Not IsMissing(StartFolder) Then
            m_StartFolder = StartFolder
            If Right$(m_StartFolder, 1) <> Chr$(0) Then m_StartFolder = m_StartFolder & Chr$(0)
            .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
        End If
    End With
    bValidateFailed = False
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList And Not bValidateFailed Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        lNull = InStr(sPath, vbNullChar)
        If lNull Then
            sPath = Left$(sPath, lNull - 1)
        End If
    End If
    BrowseForFolder = sPath
End Function

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function
