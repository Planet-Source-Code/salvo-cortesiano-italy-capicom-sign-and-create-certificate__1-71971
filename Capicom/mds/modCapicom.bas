Attribute VB_Name = "modCapicom"
Option Explicit

' .... Class CAPICOM
Public objCap As New clsCapicom

' .... CAPICOM Enum
Public Enum CapEncryptionAlgorithm
    CAPICOM_ENCRYPTION_ALGORITHM_3DES = 0
    CAPICOM_ENCRYPTION_ALGORITHM_AES = 1
    CAPICOM_ENCRYPTION_ALGORITHM_DES = 2
    CAPICOM_ENCRYPTION_ALGORITHM_RC2 = 3
    CAPICOM_ENCRYPTION_ALGORITHM_RC4 = 4
End Enum

Public Enum CapEncryptionLength
    CAPICOM_ENCRYPTION_KEY_LENGTH_128_BITS = 0
    CAPICOM_ENCRYPTION_KEY_LENGTH_192_BITS = 1
    CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS = 2
    CAPICOM_ENCRYPTION_KEY_LENGTH_40_BITS = 3
    CAPICOM_ENCRYPTION_KEY_LENGTH_56_BITS = 4
    CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM = 5
End Enum

Public Enum CapEncryptionBase
    CAPICOM_ENCODE_BASE64 = 0
    CAPICOM_ENCODE_ANY = 1
    CAPICOM_ENCODE_BINARY = 2
End Enum

Public Enum CapHashedAlgorithm
    CAPICOM_HASH_ALGORITHM_MD2 = 0
    CAPICOM_HASH_ALGORITHM_MD4 = 1
    CAPICOM_HASH_ALGORITHM_MD5 = 2
    CAPICOM_HASH_ALGORITHM_SHA_256 = 3
    CAPICOM_HASH_ALGORITHM_SHA_384 = 4
    CAPICOM_HASH_ALGORITHM_SHA_512 = 5
    CAPICOM_HASH_ALGORITHM_SHA1 = 6
End Enum
Public Function EncryptFile(ByVal sFilePath As String, ByVal sPassword As String, ByVal sTextToEncrypt As String, _
                                    Optional sIncludeTag As Boolean = False, _
                                    Optional EncryptionAlgorithm As CapEncryptionAlgorithm = CAPICOM_ENCRYPTION_ALGORITHM_AES, _
                                    Optional EncryptionLength As CapEncryptionLength = CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM, _
                                    Optional EncryptionBase As CapEncryptionBase = CAPICOM_ENCODE_BASE64) As String
    Dim sContent As String
    Dim sTemp As String
    Dim intFile As Integer
    On Local Error GoTo ErrorEncrypt
    Set objCap = New clsCapicom
    intFile = FreeFile()
    
    ' .... Ecrypt String
    sContent = objCap.Encrypt(sTextToEncrypt, sPassword, EncryptionAlgorithm, EncryptionLength, EncryptionBase)
    
    Open sFilePath For Output As #intFile
    If sIncludeTag Then Print #intFile, "-----BEGIN CERTIFICATE-----"
    ' .... Write the Result into File.
    ' .... The Function Mid$ delete the last char because contains a space!
    Print #intFile, Mid$(sContent, 1, Len(sContent) - 1)
    If sIncludeTag Then Print #intFile, "-----END CERTIFICATE-----"
    Close #intFile
    
    If sIncludeTag Then
        sTemp = "-----BEGIN CERTIFICATE-----" & vbCr & sContent & vbCr & "-----END CERTIFICATE-----"
        ' .... Returned Ecrypted String
        EncryptFile = sTemp
    Else
        ' .... Returned Ecrypted String
        EncryptFile = sContent
    End If
    
    Set objCap = Nothing
Exit Function
ErrorEncrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function DecryptFileFromFile(ByVal sFilePath As String, ByVal sPassword As String) As String
    Dim sData As String
    Dim intFile As Integer
    On Local Error GoTo ErrorDecrypt
    Set objCap = New clsCapicom
    intFile = FreeFile()
    
    ' .... Open the File to be Decrypted
    Open sFilePath For Binary Access Read Lock Read Write As #intFile
        sData = Input(LOF(intFile), intFile)
    Close #intFile
    
    ' .... Returned decrypted String
    DecryptFileFromFile = objCap.Decrypt(sData, sPassword)
Exit Function
ErrorDecrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function DecryptFromString(ByVal sContents As String, ByVal sPassword As String)
    Dim sData As String
    On Local Error GoTo ErrorDecrypt
    Set objCap = New clsCapicom
    sData = sContents
    ' .... Returned decrypted String
    DecryptFromString = objCap.Decrypt(sData, sPassword)
Exit Function
ErrorDecrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function EncryptString(ByVal sPassword As String, ByVal sTextToEncrypt As String, _
                                    Optional EncryptionAlgorithm As CapEncryptionAlgorithm = CAPICOM_ENCRYPTION_ALGORITHM_AES, _
                                    Optional EncryptionLength As CapEncryptionLength = CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM, _
                                    Optional EncryptionBase As CapEncryptionBase = CAPICOM_ENCODE_BASE64) As String
    Dim sContent As String
    On Local Error GoTo ErrorEncrypt
    Set objCap = New clsCapicom
    
    ' .... Ecrypt String
    sContent = objCap.Encrypt(sTextToEncrypt, sPassword, EncryptionAlgorithm, EncryptionLength, EncryptionBase)
    
    ' .... Returned Ecrypted String
    EncryptString = sContent
    
    Set objCap = Nothing
Exit Function
ErrorEncrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function GetCRC32(sFileName As String) As String
    Dim cStream As New clsBinaryStream
    Dim cCRC32 As New clsCRC32
    Dim lCRC32 As Long
    On Local Error GoTo ErrorCRC32
    cStream.File = sFileName
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    GetCRC32 = Hex(lCRC32)
    Set cStream = Nothing
    Set cCRC32 = Nothing
Exit Function
ErrorCRC32:
    GetCRC32 = "Error #" & Err.Number
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function SaveToFileEncrypted(sFileName As String, sPassword As String, Optional sDelimiter As String = "|") As Boolean
    Dim sString As String
    Dim sPath As String
    Dim LinesFromFile As String
    Dim NextLine As String
    Dim strTemp As String
    Dim FL As Long
    On Local Error GoTo ErrorEncrypted
    If sDelimiter = Empty Then sDelimiter = "|"
    ' .... Format string
    sString = ""
    sString = sString & "Signatory" & frmMain.ts(0).Text & sDelimiter
    sString = sString & "Certifier" & frmMain.ts(1).Text & sDelimiter
    sString = sString & "CF" & frmMain.ts(2).Text & sDelimiter
    sString = sString & "State" & frmMain.ts(4).Text & sDelimiter
    sString = sString & "SignatureID" & frmMain.ts(5).Text & sDelimiter
    sString = sString & "ValidityDate" & frmMain.ts(6).Text & sDelimiter
    sString = sString & "Name" & frmMain.ts(3).Text & sDelimiter
    sString = sString & "Note" & frmMain.ts(7).Text
    If sString = "" Then
            MsgBox "Enter information into textbox!", vbExclamation, App.Title
        Exit Function
    End If
    frmMain.txtDataKey.Text = EncryptFile(sFileName, sPassword, sString)
    sPath = GetFilePath(sFileName, Only_Path)
    strTemp = Empty
    FL = FreeFile
    Open sFileName For Input As FL
        Do Until EOF(FL)
            Line Input #FL, NextLine
            If NextLine <> Empty Then strTemp = strTemp & NextLine
        Loop
    Close FL
    Open sPath + "\reg_value_Key.reg" For Output As #2
        Print #2, "Windows Registry Editor Version 5.00"
        Print #2, "[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\" & App.EXEName & "]"
        Print #2, """Key""" & "=" & """" & strTemp & """"
    Close #2
    strTemp = Empty
    SaveToFileEncrypted = True
Exit Function
ErrorEncrypted:
        SaveToFileEncrypted = False
    Err.Clear
End Function
