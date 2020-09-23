VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Chiudi"
      Height          =   525
      Left            =   3150
      TabIndex        =   1
      Top             =   4890
      Width           =   2040
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   0
      Width           =   705
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

