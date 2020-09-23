VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngProviderType As Long
Private mstrImageName As String
Private mstrName As String

Public Property Let ImageName(ByVal Val As String)

mstrImageName = Val

End Property

Public Property Get ImageName() As String

ImageName = mstrImageName

End Property


Public Property Let ProviderType(ByVal Val As Long)

mlngProviderType = Val

End Property

Public Property Get ProviderType() As Long

ProviderType = mlngProviderType

End Property

Public Property Let Name(ByVal Val As String)

mstrName = Val

End Property


Public Property Get Name() As String

Name = mstrName

End Property

