Attribute VB_Name = "ModMain"

Option Explicit

'//// Per i controlli in stile XP
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public ris As String
Public COGNOME As String
Public NOME As String
Public SEX As String
Public DATANASCITA As String
Public CODCOM As String
Private Sub InitControlsCtx()
    On Error Resume Next
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
    On Error GoTo 0
End Sub

Public Sub Main()
'//// Istanza doppia
    If App.PrevInstance Then
        MsgBox "Applicazione già in uso! Chiuderla e ritentare.", vbExclamation, App.Title
        AppActivate App.Title
                SendKeys "+", True
            End
        Exit Sub
    End If
    
    '//// Inizializzo i controlli stile XP
    Call InitControlsCtx
    
    '//// Visualizzo il Form principale
    Load frmMain
    frmMain.Show
End Sub
