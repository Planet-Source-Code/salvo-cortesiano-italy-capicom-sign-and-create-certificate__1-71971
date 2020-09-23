VERSION 5.00
Begin VB.Form frmCF 
   BorderStyle     =   0  'None
   Caption         =   "Codice Fiscale"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frmCF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCF.frx":2CFA
   ScaleHeight     =   3555
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   4035
      TabIndex        =   7
      Top             =   3255
      Width           =   1200
   End
   Begin VB.Label lblProvincia 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   1155
      TabIndex        =   6
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Label lblCittà 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   1155
      TabIndex        =   5
      Top             =   2805
      Width           =   3975
   End
   Begin VB.Label lblSesso 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   4590
      TabIndex        =   4
      Top             =   2460
      Width           =   555
   End
   Begin VB.Label lblNome 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   1155
      TabIndex        =   3
      Top             =   2460
      Width           =   2715
   End
   Begin VB.Label lblCognome 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   1155
      TabIndex        =   2
      Top             =   2085
      Width           =   3465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Fiscale"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   3600
   End
   Begin VB.Label lblCF 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1020
      TabIndex        =   0
      Top             =   1410
      Width           =   4140
   End
End
Attribute VB_Name = "frmCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblCF.Caption = frmMain.TxtCodiceFiscale.Text
    lblSesso.Caption = SEX
    lblCittà.Caption = frmMain.txtComune.Text
    lblCognome.Caption = frmMain.txtCognome.Text
    lblNome.Caption = frmMain.txtNome.Text
    lblProvincia.Caption = frmMain.txtProvincia.Text
    lblData.Caption = frmMain.txtDataNascita.Text
    frmMain.Visible = False
    Load frmtest
    frmtest.Visible = False
    frmtest.Picture1.Picture = LoadPicture("")
    DoEvents
    Me.Visible = True
    AppActivate App.Title
    SendKeys "+", True
    Set frmtest.Picture1.Picture = CaptureForm(Me)
    frmMain.Visible = True
    ' STAMPO
    PrintPictureToFitPage Printer, frmtest.Picture1.Picture
    Printer.EndDoc
    On Error Resume Next
    Unload Me
    Unload frmtest
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub


Private Sub lblCF_Click()
    Unload Me
End Sub


Private Sub lblCittà_Click()
    Unload Me
End Sub

Private Sub lblCognome_Click()
    Unload Me
End Sub


Private Sub lblData_Click()
    Unload Me
End Sub

Private Sub lblNome_Click()
    Unload Me
End Sub

Private Sub lblProvincia_Click()
    Unload Me
End Sub

Private Sub lblSesso_Click()
    Unload Me
End Sub


