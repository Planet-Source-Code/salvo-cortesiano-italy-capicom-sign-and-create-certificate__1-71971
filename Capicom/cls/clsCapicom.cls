VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapicom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' .... Class CAPICOM 2008/2009 © Salvo Cortesiano
' ....
'*/ .... About CAPICOM:
' .... One of the primary objectives during the development of CAPICOM was to make it as easy as possible certain
' .... cryptographic operations. To affix a digital signature, leaving out the necessary checks on errors, would be
' .... sufficient only three lines:
'
' .... Dim signed As New SignedData
' .... signed.Content = bufferToSign
' .... MsgBox signed.Sign…
'\*
' .... CAPICOM control DLL, is freely redistributable, and supported by Windows 95/98 with Internet Explorer 5 or higher
' .... Windows ME, Windows NT 4 with Service Pack 4 or higher, Windows 2000 and Windows XP/Vista.

Option Explicit

' .... Private Var CAPICOM
Private m_EncryptionLength As CapEncryptionLength
Private m_EncryptionMode As CapEncryptionAlgorithm
Private m_EncryptionBase As CapEncryptionBase

' .... Private Var HASH
Private m_CapHashedAlgorithm As CapHashedAlgorithm
Public Property Get EncryptionAlgorithm() As CapEncryptionAlgorithm
    EncryptionAlgorithm = m_EncryptionMode
End Property

Public Property Let EncryptionAlgorithm(ByVal NewEncryption As CapEncryptionAlgorithm)
    m_EncryptionMode = NewEncryption
End Property

Public Property Get EncryptionLength() As CapEncryptionLength
    EncryptionLength = m_EncryptionLength
End Property

Public Property Let EncryptionLength(ByVal NewEncryptionLength As CapEncryptionLength)
    m_EncryptionLength = NewEncryptionLength
End Property

Public Property Get EncryptionBase() As CapEncryptionBase
    EncryptionBase = m_EncryptionBase
End Property

Public Property Let EncryptionBase(ByVal NewEncryptionBase As CapEncryptionBase)
    m_EncryptionBase = NewEncryptionBase
End Property

Public Property Get HashedAlgorithm() As CapHashedAlgorithm
    HashedAlgorithm = m_CapHashedAlgorithm
End Property

Public Property Let HashedAlgorithm(ByVal NewHashedAlgorithm As CapHashedAlgorithm)
    m_CapHashedAlgorithm = NewHashedAlgorithm
End Property

Public Function GetHash(ByVal strString As String, Optional HashMode As CapHashedAlgorithm = CAPICOM_HASH_ALGORITHM_MD5) As String
    Dim HashMD5 As New CAPICOM.HashedData
    On Local Error GoTo HashError
    ' .... Init Hash Algorithm
    ' .... I preferred this method, because:
    ' .... CAPICOM_HASH_ALGORITHM_SHA_384
    ' .... CAPICOM_HASH_ALGORITHM_SHA_512
    '.... Not supported on Windows XP or 2000
    Select Case HashMode
        Case CAPICOM_HASH_ALGORITHM_MD2
            HashMD5.Algorithm = CAPICOM_HASH_ALGORITHM.CAPICOM_HASH_ALGORITHM_MD2
            HashMD5.Hash (strString)
        Case CAPICOM_HASH_ALGORITHM_MD4
            HashMD5.Algorithm = CAPICOM_HASH_ALGORITHM.CAPICOM_HASH_ALGORITHM_MD4
            HashMD5.Hash (strString)
        Case CAPICOM_HASH_ALGORITHM_MD5
            HashMD5.Algorithm = CAPICOM_HASH_ALGORITHM.CAPICOM_HASH_ALGORITHM_MD5
            HashMD5.Hash (strString)
        Case CAPICOM_HASH_ALGORITHM_SHA_256
            ' .... CAPICOM_HASH_ALGORITHM_SHA_256
            ' .... I use the function StrConv() in vbFromUnicode
            ' .... to convert the String with CAPICOM_HASH_ALGORITHM_SHA_256
            ' .... Switches to native Unicode
            HashMD5.Algorithm = CAPICOM_HASH_ALGORITHM_SHA_256
            HashMD5.Hash (StrConv(strString, vbFromUnicode))
        Case CAPICOM_HASH_ALGORITHM_SHA_384
            MsgBox "HASH_ALGORITHM_SHA_384, not supported on Windows XP or 2000!", vbExclamation, App.Title
            Exit Function
        Case CAPICOM_HASH_ALGORITHM_SHA_512
            MsgBox "HASH_ALGORITHM_SHA_512, not supported on Windows XP or 2000!", vbExclamation, App.Title
            Exit Function
        Case CAPICOM_HASH_ALGORITHM_SHA1
            HashMD5.Algorithm = CAPICOM_HASH_ALGORITHM.CAPICOM_HASH_ALGORITHM_SHA1
            HashMD5.Hash (strString)
    End Select
    
    ' .... Get the Hash
    'HashMD5.Hash (strString)
    ' .... GetHash contains the HASH String
    GetHash = HashMD5.value
    ' .... Release the Hash
    Set HashMD5 = Nothing
Exit Function
HashError:
    MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    GetHash = "Error {" & Err.Number & "}"
    Set HashMD5 = Nothing
    Err.Clear
End Function

Public Function Decrypt(ByVal sMessage As String, ByVal sPassword As String) As String
    Dim DecryptData As New EncryptedData
    On Local Error GoTo ErrorDecrypt
    Call DecryptData.SetSecret(sPassword)
    Call DecryptData.Decrypt(sMessage)
    ' .... Return decrypted message
    Decrypt = DecryptData.Content
    Set DecryptData = Nothing
Exit Function
ErrorDecrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function Encrypt(ByVal sMessage As String, ByVal sPassword As String, _
            Optional EncryptionAlgorithm As CapEncryptionAlgorithm = CAPICOM_ENCRYPTION_ALGORITHM_AES, _
            Optional EncryptionLength As CapEncryptionLength = CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM, _
            Optional EncryptionBase As CapEncryptionBase = CAPICOM_ENCODE_BASE64) As String
    Dim EncryptData As New EncryptedData
    On Local Error GoTo ErrorEncrypt
    
    ' .... Init the CAPICOM to Encrypt
    EncryptData.Content = sMessage                          ' .... The String to Encrypt
    EncryptData.SetSecret sPassword                         ' .... The PassWord
    EncryptData.Algorithm.Name = EncryptionAlgorithm        ' .... The Algorithm
    
    ' NOTE:
    ' .... Disable this because return a Error= Platform not supported this Function
    ' .... The {EncryptData.Algorithm} is NOT necessary because the Length
    ' .... returned by CAPICOM_ENCODE_BASE64/CAPICOM_ENCODE_ANY/CAPICOM_ENCODE_BINARY ;)
    'EncryptData.Algorithm.KeyLength = EncryptionLength      ' .... The Length
    
    ' .... Return Encrypted string
    Encrypt = EncryptData.Encrypt(EncryptionBase)
    Set EncryptData = Nothing
Exit Function
ErrorEncrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function
