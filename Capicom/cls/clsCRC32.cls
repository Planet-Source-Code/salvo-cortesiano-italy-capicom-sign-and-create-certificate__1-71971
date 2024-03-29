VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private crc32Table() As Long
Private Const BUFFER_SIZE As Long = 8192

Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
   
   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF
      
   Dim i As Integer
   Dim iLookup As Integer
   
   For i = LBound(buffer) To UBound(buffer)
      iLookup = (crc32Result And &HFF) Xor buffer(i)
      crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
      crc32Result = crc32Result Xor crc32Table(iLookup)
   Next i
   
   GetByteArrayCrc32 = Not (crc32Result)

End Function

Public Function GetFileCrc32(ByRef stream As clsBinaryStream) As Long

   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF

   Dim buffer(0 To BUFFER_SIZE - 1) As Byte
   Dim readSize As Long
   readSize = BUFFER_SIZE

   Dim count As Integer
   count = stream.Read(buffer, readSize)
   
   Dim i As Integer
   Dim iLookup As Integer
   Dim tot As Integer
   
   Do While (count > 0)
      For i = 0 To count - 1
         iLookup = (crc32Result And &HFF) Xor buffer(i)
         crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
         crc32Result = crc32Result Xor crc32Table(iLookup)
      Next i
      count = stream.Read(buffer, readSize)
   Loop

   GetFileCrc32 = Not (crc32Result)

End Function

Private Sub Class_initialize()

    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i

End Sub
