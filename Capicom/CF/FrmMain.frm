VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Calcolo del Codice Fiscale v2.0.15"
   ClientHeight    =   6555
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   5340
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Stampa"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1935
      TabIndex        =   18
      Top             =   6045
      Width           =   1530
   End
   Begin VB.TextBox txtCodice 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4350
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   17
      Top             =   2115
      Width           =   900
   End
   Begin VB.TextBox txtComune 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1155
      MaxLength       =   40
      TabIndex        =   16
      Top             =   2820
      Width           =   2745
   End
   Begin VB.TextBox txtProvincia 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1065
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   15
      Top             =   3240
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3945
      ScaleHeight     =   450
      ScaleWidth      =   1320
      TabIndex        =   12
      Top             =   2715
      Width           =   1320
      Begin VB.OptionButton OptSessoF 
         BackColor       =   &H8000000E&
         Caption         =   "F"
         Height          =   255
         Left            =   675
         TabIndex        =   14
         Top             =   90
         Width           =   540
      End
      Begin VB.OptionButton OptSessoM 
         BackColor       =   &H8000000E&
         Caption         =   "M"
         Height          =   255
         Left            =   45
         TabIndex        =   13
         Top             =   90
         Width           =   570
      End
   End
   Begin VB.TextBox txtDataNascita 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3990
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3240
      Width           =   1245
   End
   Begin VB.TextBox txtNome 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1155
      TabIndex        =   10
      Top             =   2490
      Width           =   2715
   End
   Begin VB.TextBox txtCognome 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1155
      TabIndex        =   9
      Top             =   2115
      Width           =   2775
   End
   Begin VB.CommandButton cmdCalcola 
      Caption         =   "&Calcola"
      Enabled         =   0   'False
      Height          =   390
      Left            =   150
      TabIndex        =   8
      Top             =   6045
      Width           =   1545
   End
   Begin VB.TextBox TxtCodiceFiscale 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   990
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   7
      Top             =   1440
      Width           =   4200
   End
   Begin VB.Frame Frame4 
      Height          =   2220
      Left            =   45
      TabIndex        =   5
      Top             =   3645
      Width           =   5265
      Begin VB.Label Label1 
         Caption         =   $"FrmMain.frx":2CFA
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   75
         TabIndex        =   6
         Top             =   165
         Width           =   5115
      End
   End
   Begin VB.CommandButton EndBtn 
      Caption         =   "&Chiudi"
      Height          =   390
      Left            =   3630
      TabIndex        =   0
      Top             =   6045
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Salvo Cortesiano © 2007/08"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   1245
      TabIndex        =   4
      Top             =   900
      Width           =   3885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "v2.0.15"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   3855
      TabIndex        =   3
      Top             =   495
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Fiscale"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   1455
      TabIndex        =   2
      Top             =   345
      Width           =   2250
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Codice Fiscale"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   450
      Left            =   1425
      TabIndex        =   1
      Top             =   375
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   15
      Picture         =   "FrmMain.frx":2E89
      Top             =   30
      Width           =   5310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Dim Db As Database
    Dim Comuni As Recordset
Private Sub cmdCalcola_Click()
    On Error GoTo ErrorHandler
    COGNOME = Trim(txtCognome.Text)
    If Len(COGNOME) = 0 Then
        MsgBox "CONTROLLARE I DATI IN INPUT(COGNOME)", vbExclamation
        Exit Sub
    End If
    NOME = Trim(txtNome.Text)
    If Len(NOME) = 0 Then
        MsgBox "CONTROLLARE I DATI IN INPUT(NOME)", vbExclamation
        Exit Sub
    End If
    SEX = IIf(OptSessoM, "M", "F")
    DATANASCITA = Trim(Format(txtDataNascita.Text, "DD/MM/YYYY"))
    If Len(DATANASCITA) = 0 Or Not IsDate(DATANASCITA) Then
        MsgBox "CONTROLLARE I DATI IN INPUT(DATA NASCITA)", vbExclamation
        Exit Sub
    End If
    CODCOM = txtCodice.Text
    If Len(CODCOM) = 0 Then
        MsgBox "CONTROLLARE I DATI IN INPUT(COMUNE NASCITA)", vbExclamation
        Exit Sub
    End If
    ris = CalcolaCodiceFiscale(COGNOME, NOME, SEX, DATANASCITA, CODCOM)
    If Len(ris) <> 0 Then
        TxtCodiceFiscale.Text = ris
        TxtCodiceFiscale.Visible = True
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
        MsgBox "ERRORE NEL CALCOLO DEL CODICE FISCALE", vbExclamation, App.Title
    End If
    
Exit Sub
ErrorHandler:
    MsgBox "Errore: #" & Err.Number & "." & Chr$(13) & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    Load frmCF
    frmCF.Show
End Sub


Private Sub EndBtn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Set Db = OpenDatabase(App.Path + "\comuni.mdb", False, False)
    Set Comuni = Db.OpenRecordset("Comuni")
    Comuni.Index = "COMUNI2L"

    Me.OptSessoM.Value = True
    TxtCodiceFiscale.Visible = False
   Exit Sub
ErrorHandler:
        MsgBox "Errore: #" & Err.Number & "." & Chr$(13) & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Db.Close
    Comuni.Close
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set Db = Nothing
    Set Comuni = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
    If Dir(App.Path + "\COMUNI.ldb") <> "" Then Kill App.Path + "\COMUNI.ldb"
    End
Exit Sub
ErrorHandler:
        End
    Err.Clear
End Sub


Private Sub txtCognome_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case vbKeyBack
            
        Case vbKeySpace
            
        Case 65 To 90
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCognome_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Len(txtCognome.Text) > 0 Then cmdCalcola.Enabled = True Else cmdCalcola.Enabled = False
End Sub


Private Sub txtComune_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Len(txtComune.Text) > 0 Then cmdCalcola.Enabled = True Else cmdCalcola.Enabled = False
End Sub


Private Sub txtComune_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim S As String
    Dim TmpStr As String
    Dim Colore As Long
    On Error GoTo ErrorHandler
    Colore = 0
    TmpStr = Trim(txtComune.Text)
    If TmpStr = "" Then
        txtComune.Tag = TmpStr
        Exit Sub
    End If
    If Len(TmpStr) = 1 Then S = TmpStr
    Comuni.Seek ">=", TmpStr
    If Not Comuni.EOF And Not Comuni.NoMatch Then
      If UCase(TmpStr) = UCase(Mid(Comuni!COMU_DESCR, 1, Len(TmpStr))) Then
        If Len(TmpStr) > Len(txtComune.Tag) Then
            txtComune.Text = Comuni!COMU_DESCR
            txtComune.SelStart = Len(TmpStr)
            txtComune.SelLength = Len(txtComune.Text) - (Len(TmpStr))
            txtProvincia.Text = Comuni!COMU_PROV
            txtCodice.Text = Comuni!COMU_COD
        End If
      Else
        If txtComune.ForeColor = 0 Then MsgBox "Comune non in elenco.", vbExclamation, App.Title
        Colore = &H80&
        txtProvincia.Text = "?"
        txtCodice.Text = "?"
        txtComune.SetFocus
      End If
    End If
    If txtComune.ForeColor <> Colore Then txtComune.ForeColor = Colore
    txtComune.Tag = TmpStr
    Exit Sub
ErrorHandler:
    MsgBox "Errore: #" & Err.Number & ". " & Chr$(13) & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub


Private Sub txtDataNascita_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Len(txtDataNascita.Text) > 0 Then cmdCalcola.Enabled = True Else cmdCalcola.Enabled = False
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
            
        Case vbKeySpace
            
        Case 65 To 90
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDataNascita_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
            
        Case vbKeyBack
            
        Case 47
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function CalcolaCodiceFiscale(ByVal COGNOME As String, ByVal NOME As String, ByVal Sesso As String, ByVal DATANASCITA As String, ByVal Comune As String) As String
    Dim Consonanti As String
    Dim Vocali As String
    Dim i As Integer
    Dim Temp As String
    CalcolaCodiceFiscale = ""
    Consonanti = ""
    Vocali = ""
    For i = 1 To Len(COGNOME)
        Temp = Trim(Mid(COGNOME, i, 1))
        If InStr("AEIOU", Temp) Then Vocali = Vocali & Temp Else Consonanti = Consonanti & Temp
    Next i
    CalcolaCodiceFiscale = Left(Consonanti & Vocali & "XXX", 3)
    Consonanti = ""
    Vocali = ""
    For i = 1 To Len(NOME)
        Temp = Trim(Mid(NOME, i, 1))
        If InStr("AEIOU", Temp) Then
                Vocali = Vocali & Temp
        Else
            Consonanti = Consonanti & Temp
        End If
    Next i
    If Len(Consonanti) > 3 Then
        Consonanti = Mid(Consonanti, 1, 1) & Mid(Consonanti, 3, 2)
    End If
    
    CalcolaCodiceFiscale = CalcolaCodiceFiscale & Left(Consonanti & Vocali & "XXX", 3)
    CalcolaCodiceFiscale = CalcolaCodiceFiscale & Mid(Year(DATANASCITA), 3, 2)
    CalcolaCodiceFiscale = CalcolaCodiceFiscale & Mid("ABCDEHLMPRST", Month(DATANASCITA), 1)
    CalcolaCodiceFiscale = CalcolaCodiceFiscale & IIf(Sesso = "M", Right("00" & CStr(Day(DATANASCITA)), 2), Right("00" & CStr(40 + Day(DATANASCITA)), 2))
    CalcolaCodiceFiscale = CalcolaCodiceFiscale & Comune
    Temp = CalcolaCarattereControllo(CalcolaCodiceFiscale)
    
    If Len(Temp) <> 0 Then
        CalcolaCodiceFiscale = CalcolaCodiceFiscale & Temp
    Else
        CalcolaCodiceFiscale = "Error!"
    End If
    
End Function

Private Function CalcolaCarattereControllo(ByVal CodiceParziale As String) As String
    Dim Valori As String
    Dim Carattere As String
    Dim numeroCTRL As Integer
    Dim i As Integer
    Dim Temp As Integer
    Valori = "01,00,05,07,09,13,15,17,19,21,02,04,18,20,11,03,06,08,12,14,16,10,22,25,24,23"
    numeroCTRL = 0
    For i = 2 To 14 Step 2
        Carattere = Mid(CodiceParziale, i, 1)
        If Carattere >= "A" Then
            Temp = Asc(Carattere) - Asc("A")
        Else
            Temp = Asc(Carattere) - Asc("0")
        End If
        numeroCTRL = numeroCTRL + CInt(Temp)
    Next i
    For i = 1 To 15 Step 2
        Carattere = Mid(CodiceParziale, i, 1)
        If Carattere >= "A" Then
            Temp = CInt(Mid(Valori, (Asc(Carattere) - Asc("A")) * 3 + 1, 2))
        Else
            Temp = CInt(Mid(Valori, (Asc(Carattere) - Asc("0")) * 3 + 1, 2))
        End If
        numeroCTRL = numeroCTRL + CInt(Temp)
    Next i
   numeroCTRL = (numeroCTRL Mod 26) + 65
   CalcolaCarattereControllo = Chr(numeroCTRL)
End Function

Private Sub txtNome_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Len(txtNome.Text) > 0 Then cmdCalcola.Enabled = True Else cmdCalcola.Enabled = False
End Sub


