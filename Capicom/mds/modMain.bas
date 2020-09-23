Attribute VB_Name = "modMain"
Option Explicit

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

' .... Init Class clsIni
Public INI As New clsINI

' .... Shell
Public Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const SW_SHOWNORMAL As Long = 1
Public Const SE_ERR_NOASSOC As Long = 31

Public MyString As String

Public Enum Extract
  [Only_Extension] = 0
  [Only_FileName_and_Extension] = 1
  [Only_FileName_no_Extension] = 2
  [Only_Path] = 3
End Enum

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long


Private Const MAX_IP = 5

Private Type IPINFO
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize  As Long
    unused1 As Integer
    unused2 As Integer
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long
    mIPInfo(MAX_IP) As IPINFO
End Type

Private Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

' ....
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Sub Main()
On Error GoTo ErrorHandler
    If App.PrevInstance = True Then End
    DoEvents
    Call InitControlsCtx
    Load frmMain
    frmMain.Show
Exit Sub
ErrorHandler:
    Err.Clear
    End
End Sub

Private Sub InitControlsCtx()
 On Local Error GoTo Init_Error
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
Init_Error:
    Err.Clear
End Sub

Public Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long)
   Dim hWndDesk As Long
   Dim success As Long
   On Error GoTo ErrorHandler
   hWndDesk = GetDesktopWindow()
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
  If success = SE_ERR_NOASSOC Then
    MsgBox "Sorry. Default Application not found!", vbExclamation, App.Title
    Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
Exit Function
ErrorHandler:
    MsgBox "Error in (RunShellExecute) #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Function

Public Function GetFilePath(ByVal FileName As String, strExtract As Extract) As String
    Select Case strExtract
        'Extract only extension of File
    Case 0
         GetFilePath = Mid$(FileName, InStrRev(FileName, ".", , vbTextCompare) + 1)
        'Extract only Filename and Extension
    Case 1
        GetFilePath = Mid$(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
        'Extract only FileName
   Case 2
        GetFilePath = StripString(Mid$(FileName, InStrRev(FileName, "\", , vbTextCompare) + 1))
        'Extract only Path
   Case 3
        GetFilePath = Mid$(FileName, 1, InStrRev(FileName, "\", , vbTextCompare) - 1)
   End Select
End Function

Private Function StripString(ByVal sString As String) As String
    Dim i As Integer
    Dim sTmp As String
    On Error Resume Next
    sTmp = Mid(sString, i + 1, Len(sString))
    For i = 1 To Len(sTmp)
      If Mid(sTmp, i, 1) = "." Then
        Exit For
    Else
        MyString = Mid(sString, i + 2, Len(sString))
    End If
Next
     StripString = Left(sTmp, i - 1)
End Function

Public Function MakeDirectory(szDirectory As String) As Boolean
Dim strFolder As String
Dim szRslt As String
On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
        MakeDirectory = False
    Err.Clear
End Function

Public Function StripLeft(strString As String, strChar As String, Optional sLeftsRight As Boolean = True) As String
  On Local Error Resume Next
  Dim i As Integer
    If sLeftsRight Then
        For i = 1 To Len(strString)
            If Mid$(strString, i, 1) = strChar Then
                    StripLeft = Mid$(strString, 1, i - 1)
                Exit For
            End If
        Next
    Else
        For i = (Len(strString)) To 1 Step -1
        If Mid$(strString, i, 1) = strChar Then
                StripLeft = Mid$(strString, i + 2, Len(strString) - i + 1)
            Exit For
        End If
    Next
End If
End Function

Public Function GetAllIpsFromThisComputer() As String
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim Listing As MIB_IPADDRTABLE
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    GetIpAddrTable bBytes(0), ret, False
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        If Trim(GetAllIpsFromThisComputer) <> "" Then
            GetAllIpsFromThisComputer = GetAllIpsFromThisComputer & vbCrLf & ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        Else
            GetAllIpsFromThisComputer = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        End If
    Next
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Sub GetAllIP()
    Dim X As Integer
    Dim IPs As String
    Dim arrIPs() As String
    On Local Error Resume Next
    IPs = GetAllIpsFromThisComputer
    arrIPs = Split(IPs, vbCrLf)
    For X = 0 To UBound(arrIPs)
        If arrIPs(X) <> "127.0.0.0" And arrIPs(X) <> "127.0.0.1" And arrIPs(X) <> Empty Then
            frmMain.txtInternalIPs.AddItem arrIPs(X)
        End If
    Next X
    If frmMain.txtInternalIPs.ListCount > 0 Then
        frmMain.txtInternalIPs.Enabled = True
        frmMain.txtInternalIPs.ListIndex = 0
        frmMain.CheckIP.Enabled = True
    Else
        frmMain.txtInternalIPs.Enabled = False
        frmMain.CheckIP.Enabled = False
    End If
End Sub

Public Function DirExists(ByVal strDirName As String) As Boolean
    On Local Error Resume Next
        DirExists = (GetAttr(strDirName) And vbDirectory) = vbDirectory
    Err.Clear
End Function

Public Function GetUser_Name() As String
    Dim sBuffer As String
    Dim lSize As Long
    On Local Error GoTo ErrorUserName
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    GetUser_Name = Left$(sBuffer, lSize)
Exit Function
ErrorUserName:
        GetUser_Name = Empty
    Err.Clear
End Function

Private Function GetComputer_Name(strName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    On Local Error Resume Next
    strName = Space$(16)
    NameSize = Len(strName)
    X = GetComputerName(strName, NameSize)
End Function

Public Function Compuer_Name() As String
    Dim PCName As String
    Dim P As Long
    On Local Error Resume Next
    P = GetComputer_Name(PCName)
    Compuer_Name = PCName
End Function
