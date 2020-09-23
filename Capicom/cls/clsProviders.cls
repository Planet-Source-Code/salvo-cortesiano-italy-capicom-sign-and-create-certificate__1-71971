VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProviders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moProviders As Collection

Public Function Add(ByVal Name As String, ByVal ProviderType As Long, ByVal ImageName As String) As clsProvider

Dim oProvider As New clsProvider

oProvider.Name = Name
oProvider.ProviderType = ProviderType
oProvider.ImageName = ImageName

moProviders.Add oProvider

Set Add = oProvider

End Function

Public Function Item(ByVal Key As Variant) As clsProvider

Set Item = moProviders.Item(Key)

End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

Set NewEnum = moProviders.[_NewEnum]

End Function

Public Sub Remove(ByVal Key As Variant)

moProviders.Remove Key

End Sub





Private Sub Class_Initialize()

Set moProviders = New Collection

End Sub


Private Sub Class_Terminate()

Set moProviders = Nothing

End Sub

