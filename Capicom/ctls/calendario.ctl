VERSION 5.00
Begin VB.UserControl Calendario 
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   2895
   ToolboxBitmap   =   "calendario.ctx":0000
   Begin VB.TextBox txtYear 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   57
      Text            =   "2009"
      Top             =   0
      Width           =   570
   End
   Begin VB.ComboBox cmbMes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "calendario.ctx":0312
      Left            =   270
      List            =   "calendario.ctx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Caption         =   "n/a"
      Height          =   210
      Left            =   0
      TabIndex        =   58
      Top             =   330
      Width           =   1890
   End
   Begin VB.Image imgAdd 
      Height          =   240
      Index           =   1
      Left            =   1920
      Picture         =   "calendario.ctx":0316
      ToolTipText     =   "Subtract Year"
      Top             =   330
      Width           =   240
   End
   Begin VB.Image imgAdd 
      Height          =   240
      Index           =   0
      Left            =   2130
      Picture         =   "calendario.ctx":08A0
      ToolTipText     =   "Add Year"
      Top             =   330
      Width           =   240
   End
   Begin VB.Image imgCalendar 
      Height          =   240
      Left            =   0
      Picture         =   "calendario.ctx":0E2A
      ToolTipText     =   "Click to Show Calendar"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   55
      ToolTipText     =   "Saturday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1695
      TabIndex        =   54
      ToolTipText     =   "Friday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1350
      TabIndex        =   53
      ToolTipText     =   "Thursday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1005
      TabIndex        =   52
      ToolTipText     =   "Wednesday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   660
      TabIndex        =   51
      ToolTipText     =   "Tuesday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   50
      ToolTipText     =   "Sunday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   49
      ToolTipText     =   "Monday"
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   48
      Left            =   2040
      TabIndex        =   48
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   47
      Left            =   1695
      TabIndex        =   47
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   46
      Left            =   1350
      TabIndex        =   46
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   45
      Left            =   1005
      TabIndex        =   45
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   44
      Left            =   660
      TabIndex        =   44
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   43
      Left            =   0
      TabIndex        =   43
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   42
      Left            =   330
      TabIndex        =   42
      Top             =   570
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   2040
      TabIndex        =   41
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   1695
      TabIndex        =   40
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   1350
      TabIndex        =   39
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   1005
      TabIndex        =   38
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   660
      TabIndex        =   37
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   0
      TabIndex        =   36
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   330
      TabIndex        =   35
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   2040
      TabIndex        =   34
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   1695
      TabIndex        =   33
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   1350
      TabIndex        =   32
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   1005
      TabIndex        =   31
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   660
      TabIndex        =   30
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   0
      TabIndex        =   29
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   330
      TabIndex        =   28
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   27
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   1695
      TabIndex        =   26
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1350
      TabIndex        =   25
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   1005
      TabIndex        =   24
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   660
      TabIndex        =   23
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   0
      TabIndex        =   22
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   330
      TabIndex        =   21
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   2040
      TabIndex        =   20
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   1695
      TabIndex        =   19
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1350
      TabIndex        =   18
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1005
      TabIndex        =   17
      ToolTipText     =   "15414"
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   660
      TabIndex        =   16
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   0
      TabIndex        =   15
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   330
      TabIndex        =   14
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   2040
      TabIndex        =   13
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1695
      TabIndex        =   12
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1350
      TabIndex        =   11
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1005
      TabIndex        =   10
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   660
      TabIndex        =   9
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   330
      TabIndex        =   7
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   6
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1695
      TabIndex        =   5
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1350
      TabIndex        =   4
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1005
      TabIndex        =   3
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   2
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   300
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   330
      TabIndex        =   0
      Top             =   855
      Width           =   300
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Hide As Boolean
Private Const m_def_Hide As Boolean = False

Private m_Longdata As String
Private Const m_def_Longdata As String = ""

Private m_ShortData As String
Private Const m_def_ShortData As String = ""

Private DateOutput As Date
Private beProcess As Boolean
Private rResize As Boolean

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Private Sub cmbMes_Click()
If cmbMes.ListIndex > 0 And Not beProcess Then
    Generate (Day(DateOutput) & "/" & cmbMes.ListIndex & "/" & Year(DateOutput))
End If
End Sub

Private Sub imgAdd_Click(Index As Integer)
    On Local Error Resume Next
    If Index = 1 Then
        txtYear.Text = txtYear.Text - 1
        If cmbMes.ListIndex > 0 Then
            Generate (Day(DateOutput) & "/" & Month(DateOutput) & "/" & txtYear)
        End If
    ElseIf Index = 0 Then
        txtYear.Text = txtYear.Text + 1
        If cmbMes.ListIndex > 0 Then
            Generate (Day(DateOutput) & "/" & Month(DateOutput) & "/" & txtYear)
        End If
    End If
End Sub

Private Sub imgCalendar_Click()
If m_Hide = False Then
        rResize = True
        UserControl.Height = 2550
        UserControl.Width = 2370
        m_Hide = True
        imgCalendar.ToolTipText = "Click to Hide Calendar"
        PropertyChanged "Hide"
    ElseIf m_Hide = True Then
        UserControl.Height = 240
        UserControl.Width = 240
        rResize = False
        m_Hide = False
        imgCalendar.ToolTipText = "Click to Show Calendar"
        PropertyChanged "Hide"
    End If
    RaiseEvent Click
End Sub

Private Sub imgCalendar_DblClick()
    RaiseEvent DblClick
End Sub


Private Sub imgCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub imgCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub imgCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub lblDay_Click(Index As Integer)
    Call RestableColors
    lblDay(Index).BackColor = &H80000008
    lblDay(Index).ForeColor = &HFFFFFF
    RaiseEvent Click
End Sub

Private Sub RestableColors()
Dim intI As Integer
For intI = 1 To 42
    If lblDay(intI).Visible Then
        If lblDay(intI) <> "" Then
            If CInt(lblDay(intI)) > 8 And intI < 7 Then
                lblDay(intI).ForeColor = &HC0C0C0
            ElseIf lblDay(intI) < 8 And intI > 25 Then
                lblDay(intI).ForeColor = &HC0C0C0
            Else
                lblDay(intI).ForeColor = &H404040
            End If
            lblDay(intI).BackColor = &HFFFFFF
        End If
    End If
Next intI
End Sub
Public Function Generate(strDateDefault As Date)
Dim IntDaysOfWeek
Dim DaysOfWeek: Dim strDay: Dim strMonth: Dim strYear: Dim strDate: Dim intI: Dim IntJ
Dim strFechaOut As Date
Dim intCantAct: Dim intDay: Dim intCantDes
If Not IsDate(strDateDefault) Then strDateDefault = Date
DateOutput = strDateDefault
RestableColors
beProcess = True
cmbMes.ListIndex = Month(strDateDefault)
beProcess = False
txtYear.Text = Year(strDateDefault)
strDay = Format(strDateDefault, "dd")
strMonth = Format(strDateDefault, "mm")
strYear = Format(strDateDefault, "yyyy")
strDate = "01/" & strMonth & "/" & strYear
DaysOfWeek = Weekday(strDate)
Select Case DaysOfWeek
    Case 1, 2, 3, 4, 5:
        IntDaysOfWeek = 5
        lblDay(36).Visible = False
        lblDay(37).Visible = False
        lblDay(38).Visible = False
        lblDay(39).Visible = False
        lblDay(40).Visible = False
        lblDay(41).Visible = False
        lblDay(42).Visible = False
    Case 6
        If strMonth = 11 Then
            IntDaysOfWeek = 5
            lblDay(36).Visible = False
            lblDay(37).Visible = False
            lblDay(38).Visible = False
            lblDay(39).Visible = False
            lblDay(40).Visible = False
            lblDay(41).Visible = False
            lblDay(42).Visible = False
        Else
            IntDaysOfWeek = 6
            lblDay(36).Visible = True
            lblDay(37).Visible = True
            lblDay(38).Visible = True
            lblDay(39).Visible = True
            lblDay(40).Visible = True
            lblDay(41).Visible = True
            lblDay(42).Visible = True
        End If
    Case Else
        IntDaysOfWeek = 6
        lblDay(36).Visible = True
        lblDay(37).Visible = True
        lblDay(38).Visible = True
        lblDay(39).Visible = True
        lblDay(40).Visible = True
        lblDay(41).Visible = True
        lblDay(42).Visible = True
End Select
'Day's
For intI = 1 To IntDaysOfWeek
    For IntJ = 1 To 7
        intCantAct = intCantAct + 1
        If intCantAct >= DaysOfWeek Then
            intDay = intDay + 1
            If IsDate(intDay & "/" & strMonth & "/" & strYear) Then
                strFechaOut = intDay & "/" & strMonth & "/" & strYear
                lblDay(intCantAct) = intDay
                lblDay(intCantAct).ForeColor = &H80000008
                lblDay(intCantAct).ToolTipText = Format(strFechaOut, "dd/mm/yyyy")
            Else
                intCantDes = intCantDes + 1
                strFechaOut = intCantDes & "/" & Month(DateAdd("m", 1, strFechaOut)) & "/" & strYear
                lblDay(intCantAct) = intCantDes
                lblDay(intCantAct).ForeColor = &HC0C0C0
                lblDay(intCantAct).ToolTipText = Format(strFechaOut, "dd/mm/yyyy")
            End If
        Else
            strFechaOut = DateAdd("d", intCantAct - DaysOfWeek, strDate)
            lblDay(intCantAct) = Day(DateAdd("d", intCantAct - DaysOfWeek, strDate))
            lblDay(intCantAct).ForeColor = &HC0C0C0
            lblDay(intCantAct).ToolTipText = Format(strFechaOut, "dd/mm/yyyy")
        End If
        If lblDay(intCantAct) = strDay Then
            lblDay(intCantAct).BackColor = &H80000008
            lblDay(intCantAct).ForeColor = &HFFFFFF
        End If
    Next IntJ
Next intI
m_Longdata = Format(DateOutput, "Long Date")
PropertyChanged "LongData"
lblDate.Caption = Format(DateOutput, "Short Date")
m_ShortData = Format(DateOutput, "Short Date")
PropertyChanged "ShortData"
End Function

Private Sub lblDay_DblClick(Index As Integer)
    Dim intI As Integer
    DateOutput = Format(lblDay(Index).ToolTipText, "dd/mm/yyyy")
    lblDate.Caption = Format(DateOutput, "Short Date")
    Call Generate(Format(lblDay(Index).ToolTipText, "dd/mm/yyyy"))
    Call RestableColors
    lblDay(Index).BackColor = &H80000008
    lblDay(Index).ForeColor = &HFFFFFF
    UserControl.Height = 240
    UserControl.Width = 240
    rResize = False
    m_Hide = False
    imgCalendar.ToolTipText = "Click to Show Calendar"
    PropertyChanged "Hide"
    RaiseEvent DblClick
End Sub

Private Sub lblDay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub lblDay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
beProcess = False
rResize = False
m_Longdata = m_def_Longdata
m_ShortData = m_def_ShortData
m_Hide = m_def_Hide
PropertyChanged "Hide"
imgCalendar.ToolTipText = "Click to Show Calendar"
'Month's
cmbMes.AddItem "Calendar"
cmbMes.AddItem "January"
cmbMes.AddItem "February"
cmbMes.AddItem "March"
cmbMes.AddItem "April"
cmbMes.AddItem "May"
cmbMes.AddItem "June"
cmbMes.AddItem "July"
cmbMes.AddItem "August"
cmbMes.AddItem "September"
cmbMes.AddItem "October"
cmbMes.AddItem "November"
cmbMes.AddItem "December"
Call UserControl_Resize
Call Generate(Now)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Longdata = PropBag.ReadProperty("LongData", m_def_Longdata)
    m_ShortData = PropBag.ReadProperty("ShortData", m_def_ShortData)
    m_Hide = PropBag.ReadProperty("Hide", m_def_Hide)
End Sub


Private Sub UserControl_Resize()
    If rResize = False Then
        UserControl.Height = 240
        UserControl.Width = 240
    End If
End Sub

Public Property Get LongData() As String
    LongData = m_Longdata
End Property

Public Property Let LongData(ByVal NewValue As String)
    m_Longdata = NewValue
    PropertyChanged "LongData"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("LongData", m_Longdata, m_def_Longdata)
    Call PropBag.WriteProperty("ShortData", m_ShortData, m_def_ShortData)
    Call PropBag.WriteProperty("Hide", m_Hide, m_def_Hide)
End Sub
Public Property Get Hide() As Boolean
    Hide = m_Hide
End Property

Public Property Let Hide(ByVal NewValue As Boolean)
    m_Hide = NewValue
    If m_Hide = False Then
        rResize = True
        UserControl.Height = 540
        UserControl.Width = 2340
    ElseIf m_Hide = True Then
        UserControl.Height = 2880
        UserControl.Width = 2340
        rResize = False
    End If
    PropertyChanged "Hide"
End Property

Public Property Get ShortData() As String
    ShortData = m_ShortData
End Property

Public Property Let ShortData(ByVal NewData As String)
    m_ShortData = NewData
    lblDate.Caption = Format(DateOutput, "Short Date")
    PropertyChanged "ShortData"
End Property
