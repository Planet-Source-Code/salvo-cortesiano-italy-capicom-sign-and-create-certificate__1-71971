VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "n/a"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   13560
   Begin VB.TextBox txtInfo 
      Height          =   1725
      Left            =   5385
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   146
      Text            =   "frmMain.frx":23D2
      Top             =   5895
      Width           =   8085
   End
   Begin VB.Frame Frame2 
      Caption         =   "Algorithm Option"
      Height          =   1905
      Left            =   60
      TabIndex        =   22
      Top             =   5775
      Width           =   5235
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   30
         ScaleHeight     =   1665
         ScaleWidth      =   5160
         TabIndex        =   23
         Top             =   195
         Width           =   5160
         Begin VB.ComboBox cmbBase 
            Height          =   330
            ItemData        =   "frmMain.frx":3B66
            Left            =   1305
            List            =   "frmMain.frx":3B73
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1305
            Width           =   3690
         End
         Begin VB.ComboBox cmbLength 
            Height          =   330
            ItemData        =   "frmMain.frx":3BA1
            Left            =   1305
            List            =   "frmMain.frx":3BB7
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   900
            Width           =   3690
         End
         Begin VB.ComboBox cmbEncryption 
            Height          =   330
            ItemData        =   "frmMain.frx":3C78
            Left            =   1305
            List            =   "frmMain.frx":3C8B
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   495
            Width           =   3690
         End
         Begin VB.ComboBox cmbAlgorithm 
            Height          =   330
            ItemData        =   "frmMain.frx":3D12
            Left            =   1305
            List            =   "frmMain.frx":3D2B
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   90
            Width           =   3690
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Base:"
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Length:"
            Height          =   255
            Left            =   90
            TabIndex        =   29
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Algorithm:"
            Height          =   255
            Left            =   90
            TabIndex        =   27
            Top             =   510
            Width           =   1080
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Hash:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   105
            Width           =   1080
         End
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   12795
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "CAPICOM Tools"
      Height          =   4230
      Left            =   60
      TabIndex        =   8
      Top             =   1515
      Width           =   13455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4005
         Left            =   15
         ScaleHeight     =   4005
         ScaleWidth      =   13395
         TabIndex        =   9
         Top             =   180
         Width           =   13395
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3480
            Index           =   6
            Left            =   90
            ScaleHeight     =   3480
            ScaleWidth      =   13215
            TabIndex        =   204
            Top             =   420
            Visible         =   0   'False
            Width           =   13215
            Begin VB.Frame Frame10 
               BackColor       =   &H80000009&
               Height          =   1290
               Left            =   5040
               TabIndex        =   228
               Top             =   2145
               Width           =   8130
               Begin VB.PictureBox Picture9 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   1095
                  Left            =   60
                  ScaleHeight     =   1095
                  ScaleWidth      =   8025
                  TabIndex        =   229
                  Top             =   135
                  Width           =   8025
                  Begin VB.ComboBox cmbsStoresList 
                     Height          =   330
                     ItemData        =   "frmMain.frx":3DC8
                     Left            =   5505
                     List            =   "frmMain.frx":3DD8
                     Style           =   2  'Dropdown List
                     TabIndex        =   237
                     Top             =   60
                     Width           =   2445
                  End
                  Begin VB.CheckBox CheckEncrypt 
                     BackColor       =   &H80000009&
                     Caption         =   "Encrypt Message"
                     Height          =   255
                     Left            =   3390
                     TabIndex        =   236
                     Top             =   75
                     Width           =   1980
                  End
                  Begin VB.CommandButton cmdResolveName 
                     Caption         =   "Resolve Name"
                     Height          =   360
                     Left            =   60
                     TabIndex        =   235
                     Top             =   690
                     Width           =   1665
                  End
                  Begin VB.CommandButton cmdSaveMessage 
                     Caption         =   "Save Message"
                     Height          =   360
                     Left            =   1860
                     TabIndex        =   234
                     Top             =   690
                     Width           =   1665
                  End
                  Begin VB.TextBox txtProvider 
                     Alignment       =   2  'Center
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
                     Left            =   75
                     TabIndex        =   233
                     Text            =   "hotmail.com"
                     Top             =   75
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdOpenMessage 
                     Caption         =   "Open Message *.eml"
                     Height          =   360
                     Left            =   3660
                     TabIndex        =   232
                     Top             =   690
                     Width           =   2520
                  End
                  Begin VB.CommandButton cmdNewEmail 
                     Caption         =   "New E-Mail"
                     Height          =   360
                     Left            =   6300
                     TabIndex        =   231
                     Top             =   690
                     Width           =   1605
                  End
                  Begin VB.CheckBox CheckSign 
                     BackColor       =   &H80000009&
                     Caption         =   "Sign Message"
                     Height          =   255
                     Left            =   1740
                     TabIndex        =   230
                     Top             =   75
                     Value           =   1  'Checked
                     Width           =   1665
                  End
               End
            End
            Begin VB.CommandButton btnGoodSig 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   0
               MaskColor       =   &H00FF00FF&
               Style           =   1  'Graphical
               TabIndex        =   227
               Top             =   390
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton btnGoodEnc 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   0
               MaskColor       =   &H00FF00FF&
               Style           =   1  'Graphical
               TabIndex        =   226
               Top             =   0
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Frame gboxViewMail 
               BackColor       =   &H80000009&
               BorderStyle     =   0  'None
               Height          =   1935
               Left            =   4950
               TabIndex        =   215
               Top             =   330
               Width           =   8190
               Begin VB.TextBox txtToValue 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   220
                  Top             =   480
                  Width           =   6825
               End
               Begin VB.TextBox txtFromValue 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   219
                  Top             =   120
                  Width           =   6825
               End
               Begin VB.TextBox txtDateValue 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   218
                  Top             =   840
                  Width           =   6825
               End
               Begin VB.TextBox txtCCValue 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   217
                  Top             =   1200
                  Width           =   6825
               End
               Begin VB.TextBox txtSubjectValue 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   216
                  Top             =   1560
                  Width           =   6825
               End
               Begin VB.Label lblTo 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "To:"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   225
                  Top             =   480
                  Width           =   855
               End
               Begin VB.Label lblCC 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "CC:"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   224
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.Label lblSubject 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Subject:"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   223
                  Top             =   1560
                  Width           =   855
               End
               Begin VB.Label lblDate 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date:"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   222
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label lblFrom 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "From:"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   221
                  Top             =   120
                  Width           =   855
               End
            End
            Begin VB.CommandButton btnBadSig 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   0
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmMain.frx":3DFB
               Style           =   1  'Graphical
               TabIndex        =   214
               Top             =   375
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton btnBadEnc 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   0
               MaskColor       =   &H00FF00FF&
               Picture         =   "frmMain.frx":413D
               Style           =   1  'Graphical
               TabIndex        =   213
               Top             =   15
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtMessageBody 
               Height          =   1830
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   211
               Top             =   1605
               Width           =   4815
            End
            Begin VB.TextBox txtSubject 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   207
               Top             =   1050
               Width           =   3720
            End
            Begin VB.TextBox txtCC 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   206
               Top             =   690
               Width           =   3720
            End
            Begin VB.TextBox txtTo 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1140
               TabIndex        =   205
               Top             =   330
               Width           =   3720
            End
            Begin VB.Label lblSubject 
               BackStyle       =   0  'Transparent
               Caption         =   "Body Message:"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   212
               Top             =   1380
               Width           =   1680
            End
            Begin VB.Label lblSubject 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Subject:"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   210
               Top             =   1050
               Width           =   855
            End
            Begin VB.Label lblCC 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CC:"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   209
               Top             =   690
               Width           =   855
            End
            Begin VB.Label lblTo 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   208
               Top             =   330
               Width           =   855
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            Height          =   3495
            Index           =   5
            Left            =   90
            ScaleHeight     =   3495
            ScaleWidth      =   13230
            TabIndex        =   145
            Top             =   420
            Visible         =   0   'False
            Width           =   13230
            Begin VB.Frame Frame9 
               BackColor       =   &H80000009&
               Height          =   3390
               Left            =   6015
               TabIndex        =   166
               Top             =   60
               Width           =   7170
               Begin VB.PictureBox picCab 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3120
                  Left            =   30
                  ScaleHeight     =   3120
                  ScaleWidth      =   7095
                  TabIndex        =   198
                  Top             =   225
                  Visible         =   0   'False
                  Width           =   7095
                  Begin VB.CommandButton cmdOpenFolder 
                     Caption         =   "Open Folder"
                     Height          =   345
                     Left            =   105
                     TabIndex        =   203
                     Top             =   2700
                     Width           =   1620
                  End
                  Begin VB.CommandButton cmdMakCab 
                     Caption         =   "Make CAB File"
                     Height          =   345
                     Left            =   1875
                     TabIndex        =   202
                     Top             =   2685
                     Width           =   1680
                  End
                  Begin VB.CommandButton cmdRemove 
                     Caption         =   "Remove File"
                     Height          =   345
                     Left            =   3720
                     TabIndex        =   201
                     Top             =   2685
                     Width           =   1575
                  End
                  Begin VB.CommandButton cmdAdd 
                     Caption         =   "Added Files"
                     Height          =   345
                     Left            =   5430
                     TabIndex        =   200
                     Top             =   2685
                     Width           =   1560
                  End
                  Begin VB.ListBox lstList 
                     Height          =   2580
                     Left            =   60
                     TabIndex        =   199
                     Top             =   30
                     Width           =   6990
                  End
               End
               Begin VB.PictureBox picTools 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3105
                  Left            =   30
                  ScaleHeight     =   3105
                  ScaleWidth      =   7095
                  TabIndex        =   183
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   7095
                  Begin VB.TextBox txtSHA1 
                     Height          =   315
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   188
                     Top             =   2685
                     Width           =   6900
                  End
                  Begin VB.TextBox txtBase64 
                     Height          =   2115
                     Left            =   105
                     Locked          =   -1  'True
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   185
                     Top             =   300
                     Width           =   6900
                  End
                  Begin VB.Label Label59 
                     BackColor       =   &H80000009&
                     Caption         =   "SHA1:"
                     Height          =   240
                     Left            =   105
                     TabIndex        =   187
                     Top             =   2430
                     Width           =   675
                  End
                  Begin VB.Label Label58 
                     BackColor       =   &H80000009&
                     Caption         =   "Base64:"
                     Height          =   225
                     Left            =   120
                     TabIndex        =   186
                     Top             =   60
                     Width           =   900
                  End
               End
               Begin VB.PictureBox picSetting 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3135
                  Left            =   60
                  ScaleHeight     =   3135
                  ScaleWidth      =   7065
                  TabIndex        =   167
                  Top             =   210
                  Width           =   7065
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.4.1.311.10.3.1"
                     Height          =   270
                     Index           =   6
                     Left            =   45
                     TabIndex        =   174
                     Top             =   1695
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.4.1.311.10.3.4"
                     Height          =   270
                     Index           =   5
                     Left            =   45
                     TabIndex        =   173
                     Top             =   1425
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.5.5.7.3.8"
                     Height          =   270
                     Index           =   4
                     Left            =   45
                     TabIndex        =   172
                     Top             =   1170
                     Value           =   1  'Checked
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.5.5.7.3.4"
                     Height          =   270
                     Index           =   3
                     Left            =   45
                     TabIndex        =   171
                     Top             =   900
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.5.5.7.3.3"
                     Height          =   270
                     Index           =   2
                     Left            =   45
                     TabIndex        =   170
                     Top             =   630
                     Value           =   1  'Checked
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.5.5.7.3.2"
                     Height          =   270
                     Index           =   1
                     Left            =   45
                     TabIndex        =   169
                     Top             =   360
                     Width           =   2730
                  End
                  Begin VB.CheckBox CheckUsers 
                     BackColor       =   &H80000009&
                     Caption         =   "1.3.6.1.5.5.7.3.1"
                     Height          =   270
                     Index           =   0
                     Left            =   45
                     TabIndex        =   168
                     Top             =   105
                     Width           =   2730
                  End
                  Begin VB.Label Label57 
                     BackColor       =   &H80000009&
                     Caption         =   $"frmMain.frx":447F
                     Height          =   720
                     Left            =   60
                     TabIndex        =   182
                     Top             =   2415
                     Width           =   6930
                  End
                  Begin VB.Line Line4 
                     X1              =   1335
                     X2              =   1335
                     Y1              =   2205
                     Y2              =   2370
                  End
                  Begin VB.Line Line3 
                     X1              =   2700
                     X2              =   2700
                     Y1              =   2205
                     Y2              =   1995
                  End
                  Begin VB.Line Line2 
                     X1              =   135
                     X2              =   2700
                     Y1              =   2205
                     Y2              =   2205
                  End
                  Begin VB.Line Line1 
                     X1              =   135
                     X2              =   135
                     Y1              =   2010
                     Y2              =   2205
                  End
                  Begin VB.Image Image6 
                     Height          =   480
                     Left            =   6450
                     Picture         =   "frmMain.frx":452E
                     Top             =   885
                     Width           =   480
                  End
                  Begin VB.Image Image5 
                     Height          =   480
                     Left            =   6420
                     Picture         =   "frmMain.frx":4DF8
                     Top             =   105
                     Width           =   480
                  End
                  Begin VB.Label Label56 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Microsoft Trust List Signing!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   181
                     Top             =   1680
                     Width           =   3480
                  End
                  Begin VB.Label Label55 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Encrypting File System!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   180
                     Top             =   1440
                     Width           =   2820
                  End
                  Begin VB.Label Label54 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Time Stamping!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   179
                     Top             =   1170
                     Width           =   1905
                  End
                  Begin VB.Label Label53 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Secure Email!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   178
                     Top             =   900
                     Width           =   1905
                  End
                  Begin VB.Label Label52 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Code Signing!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   177
                     Top             =   630
                     Width           =   1905
                  End
                  Begin VB.Label Label51 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Only Client!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   176
                     Top             =   360
                     Width           =   1860
                  End
                  Begin VB.Label Label50 
                     BackColor       =   &H80000009&
                     Caption         =   "<- Only Server!"
                     Height          =   285
                     Left            =   2925
                     TabIndex        =   175
                     Top             =   105
                     Width           =   1815
                  End
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H80000009&
               Caption         =   "Certificate"
               Height          =   3390
               Left            =   45
               TabIndex        =   147
               Top             =   60
               Width           =   5925
               Begin VB.PictureBox Picture8 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3135
                  Left            =   30
                  ScaleHeight     =   3135
                  ScaleWidth      =   5850
                  TabIndex        =   148
                  Top             =   210
                  Width           =   5850
                  Begin VB.CommandButton cmdCreateCab 
                     Caption         =   "Create CAB"
                     Height          =   345
                     Left            =   4275
                     TabIndex        =   197
                     Top             =   1935
                     Width           =   1440
                  End
                  Begin ProjectCapicom.Calendario CalVal 
                     Height          =   240
                     Left            =   120
                     TabIndex        =   196
                     Top             =   795
                     Width           =   240
                     _ExtentX        =   423
                     _ExtentY        =   423
                     LongData        =   "giovedì 9 aprile 2009"
                     ShortData       =   "09/04/2009"
                  End
                  Begin VB.CheckBox CheckInstall 
                     BackColor       =   &H80000009&
                     Caption         =   "Create and Install"
                     Height          =   405
                     Left            =   2790
                     TabIndex        =   195
                     Top             =   2655
                     Width           =   1395
                  End
                  Begin VB.ComboBox cmb_Stores 
                     Height          =   330
                     ItemData        =   "frmMain.frx":56C2
                     Left            =   1920
                     List            =   "frmMain.frx":56D2
                     Style           =   2  'Dropdown List
                     TabIndex        =   193
                     Top             =   2265
                     Width           =   2235
                  End
                  Begin VB.CheckBox CheckPutMyValidity 
                     BackColor       =   &H80000009&
                     Height          =   240
                     Left            =   5475
                     TabIndex        =   192
                     ToolTipText     =   "Put your Validity Date"
                     Top             =   1245
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox CheckIP 
                     BackColor       =   &H80000009&
                     Caption         =   "Include my Internal IP"
                     Height          =   240
                     Left            =   45
                     TabIndex        =   191
                     Top             =   2640
                     Width           =   2685
                  End
                  Begin VB.ComboBox txtInternalIPs 
                     Height          =   330
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   190
                     Top             =   1920
                     Width           =   2235
                  End
                  Begin VB.CommandButton cmdPicTools 
                     Caption         =   "Base64 >>"
                     Height          =   345
                     Left            =   4275
                     TabIndex        =   184
                     Top             =   2340
                     Width           =   1440
                  End
                  Begin VB.CheckBox CheckPfx 
                     BackColor       =   &H80000009&
                     Caption         =   "Create also *.pfx file"
                     Height          =   270
                     Left            =   45
                     TabIndex        =   165
                     Top             =   2850
                     Value           =   1  'Checked
                     Width           =   2685
                  End
                  Begin VB.TextBox txtPassCert 
                     Height          =   270
                     Left            =   1920
                     TabIndex        =   163
                     Top             =   1590
                     Width           =   3465
                  End
                  Begin VB.TextBox txtValidTo 
                     Height          =   270
                     Left            =   4230
                     TabIndex        =   162
                     Top             =   1215
                     Width           =   1170
                  End
                  Begin VB.TextBox txtValidFrom 
                     Height          =   270
                     Left            =   1920
                     TabIndex        =   159
                     Top             =   1215
                     Width           =   1200
                  End
                  Begin VB.TextBox txtAutority 
                     Height          =   270
                     Left            =   1920
                     TabIndex        =   157
                     Top             =   855
                     Width           =   3465
                  End
                  Begin VB.TextBox txtCertificateName 
                     Height          =   270
                     Left            =   1920
                     TabIndex        =   155
                     Top             =   480
                     Width           =   3450
                  End
                  Begin VB.CommandButton cmdCreateCertificate 
                     Caption         =   "Create"
                     Height          =   345
                     Left            =   4275
                     TabIndex        =   154
                     Top             =   2745
                     Width           =   1440
                  End
                  Begin VB.TextBox txtDestPath 
                     Height          =   270
                     Left            =   615
                     TabIndex        =   150
                     Top             =   150
                     Width           =   4530
                  End
                  Begin VB.CommandButton cmdFolderBrowse 
                     Caption         =   "..."
                     Height          =   270
                     Left            =   5250
                     TabIndex        =   149
                     ToolTipText     =   "Certificate Folder Path"
                     Top             =   135
                     Width           =   525
                  End
                  Begin VB.Label Label61 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Store:"
                     Height          =   210
                     Left            =   1080
                     TabIndex        =   194
                     Top             =   2340
                     Width           =   750
                  End
                  Begin VB.Label Label60 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "IP's and User:"
                     Height          =   210
                     Left            =   60
                     TabIndex        =   189
                     Top             =   1965
                     Width           =   1770
                  End
                  Begin VB.Image Image7 
                     Height          =   480
                     Left            =   5370
                     Picture         =   "frmMain.frx":56F5
                     Top             =   615
                     Width           =   480
                  End
                  Begin VB.Label Label49 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "PassWord:"
                     Height          =   270
                     Left            =   45
                     TabIndex        =   164
                     Top             =   1605
                     Width           =   1785
                  End
                  Begin VB.Label Label48 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Valid To:"
                     Height          =   270
                     Left            =   3135
                     TabIndex        =   161
                     Top             =   1215
                     Width           =   1065
                  End
                  Begin VB.Label Label47 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Valid From:"
                     Height          =   270
                     Left            =   45
                     TabIndex        =   160
                     Top             =   1215
                     Width           =   1785
                  End
                  Begin VB.Label Label46 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Autority:"
                     Height          =   270
                     Left            =   45
                     TabIndex        =   158
                     Top             =   855
                     Width           =   1785
                  End
                  Begin VB.Label Label45 
                     BackColor       =   &H80000009&
                     Caption         =   "Certificate Name:"
                     Height          =   285
                     Left            =   45
                     TabIndex        =   156
                     Top             =   495
                     Width           =   1830
                  End
                  Begin VB.Label Label42 
                     BackColor       =   &H80000009&
                     Caption         =   "Path:"
                     Height          =   285
                     Left            =   45
                     TabIndex        =   151
                     Top             =   150
                     Width           =   570
                  End
               End
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3495
            Index           =   4
            Left            =   90
            ScaleHeight     =   3495
            ScaleWidth      =   13230
            TabIndex        =   106
            Top             =   420
            Visible         =   0   'False
            Width           =   13230
            Begin VB.PictureBox picCrypt 
               BackColor       =   &H80000009&
               BorderStyle     =   0  'None
               Height          =   3000
               Index           =   2
               Left            =   90
               ScaleHeight     =   3000
               ScaleWidth      =   13050
               TabIndex        =   126
               Top             =   405
               Visible         =   0   'False
               Width           =   13050
               Begin VB.CommandButton cmdDecryptFileEasy 
                  Caption         =   "Decrypt File (easy)"
                  Height          =   375
                  Left            =   10215
                  TabIndex        =   144
                  Top             =   2325
                  Width           =   2640
               End
               Begin VB.CommandButton cmdDecryptFile 
                  Caption         =   "Decrypt File"
                  Height          =   375
                  Left            =   3990
                  TabIndex        =   143
                  Top             =   2325
                  Width           =   1725
               End
               Begin VB.CommandButton cmdEncryptFile 
                  Caption         =   "Encrypt File"
                  Height          =   375
                  Left            =   5835
                  TabIndex        =   142
                  Top             =   2325
                  Width           =   1710
               End
               Begin VB.CommandButton cmdEncryptFileEasy 
                  Caption         =   "Encrypt File (easy)"
                  Height          =   375
                  Left            =   7665
                  TabIndex        =   141
                  Top             =   2325
                  Width           =   2415
               End
               Begin ComctlLib.TreeView trvKeyResults 
                  Height          =   1830
                  Left            =   5805
                  TabIndex        =   139
                  Top             =   285
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   3228
                  _Version        =   327682
                  Style           =   7
                  Appearance      =   1
               End
               Begin VB.CommandButton cmdCreate 
                  Caption         =   "Create Key"
                  Height          =   375
                  Left            =   2220
                  TabIndex        =   138
                  Top             =   2325
                  Width           =   1575
               End
               Begin VB.CheckBox chkExportKey 
                  BackColor       =   &H8000000E&
                  Caption         =   "Exportable key"
                  Height          =   255
                  Left            =   105
                  TabIndex        =   137
                  Top             =   2445
                  Width           =   2025
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H80000009&
                  Caption         =   "Key Generation"
                  Height          =   1470
                  Left            =   75
                  TabIndex        =   129
                  Top             =   645
                  Width           =   5610
                  Begin VB.PictureBox Picture6 
                     BackColor       =   &H80000009&
                     BorderStyle     =   0  'None
                     Height          =   1230
                     Left            =   30
                     ScaleHeight     =   1230
                     ScaleWidth      =   5535
                     TabIndex        =   130
                     Top             =   195
                     Width           =   5535
                     Begin VB.PictureBox Picture7 
                        BackColor       =   &H80000009&
                        BorderStyle     =   0  'None
                        Height          =   825
                        Left            =   45
                        ScaleHeight     =   825
                        ScaleWidth      =   5445
                        TabIndex        =   133
                        Top             =   360
                        Width           =   5445
                        Begin VB.OptionButton optSalt 
                           BackColor       =   &H80000009&
                           Caption         =   "Create salt"
                           Height          =   255
                           Index           =   0
                           Left            =   105
                           TabIndex        =   135
                           Top             =   375
                           Width           =   1680
                        End
                        Begin VB.OptionButton optSalt 
                           BackColor       =   &H80000009&
                           Caption         =   "Don't add salt"
                           Height          =   255
                           Index           =   1
                           Left            =   1845
                           TabIndex        =   134
                           Top             =   375
                           Width           =   3555
                        End
                        Begin VB.Label lblMain 
                           BackColor       =   &H80000009&
                           Caption         =   "Salt Option:"
                           ForeColor       =   &H8000000D&
                           Height          =   255
                           Index           =   0
                           Left            =   105
                           TabIndex        =   136
                           Top             =   75
                           Width           =   1545
                        End
                     End
                     Begin VB.OptionButton optKeyGen 
                        BackColor       =   &H80000009&
                        Caption         =   "Random Key"
                        Height          =   255
                        Index           =   0
                        Left            =   105
                        TabIndex        =   132
                        Top             =   105
                        Width           =   1785
                     End
                     Begin VB.OptionButton optKeyGen 
                        BackColor       =   &H80000009&
                        Caption         =   "Use Hash material Key"
                        Height          =   255
                        Index           =   1
                        Left            =   1950
                        TabIndex        =   131
                        Top             =   120
                        Width           =   3540
                     End
                  End
               End
               Begin VB.ComboBox cboKeyAlgs 
                  Height          =   330
                  ItemData        =   "frmMain.frx":5FBF
                  Left            =   60
                  List            =   "frmMain.frx":5FC1
                  Style           =   2  'Dropdown List
                  TabIndex        =   127
                  Top             =   285
                  Width           =   5655
               End
               Begin VB.Label lblMain 
                  BackColor       =   &H80000009&
                  Caption         =   "Result:"
                  Height          =   255
                  Index           =   1
                  Left            =   5805
                  TabIndex        =   140
                  Top             =   30
                  Width           =   990
               End
               Begin VB.Label lblMain 
                  BackColor       =   &H80000009&
                  Caption         =   "Keys:"
                  Height          =   255
                  Index           =   5
                  Left            =   75
                  TabIndex        =   128
                  Top             =   45
                  Width           =   735
               End
            End
            Begin VB.PictureBox picCrypt 
               BackColor       =   &H80000009&
               BorderStyle     =   0  'None
               Height          =   3000
               Index           =   1
               Left            =   90
               ScaleHeight     =   3000
               ScaleWidth      =   13050
               TabIndex        =   116
               Top             =   405
               Visible         =   0   'False
               Width           =   13050
               Begin VB.CommandButton cmdCreateHash 
                  Caption         =   "Create Hash"
                  Height          =   375
                  Left            =   6135
                  TabIndex        =   124
                  Top             =   780
                  Width           =   1620
               End
               Begin VB.CommandButton cmdSignAndVerify 
                  Caption         =   "Sign and Verify"
                  Height          =   375
                  Left            =   7875
                  TabIndex        =   123
                  Top             =   780
                  Width           =   1950
               End
               Begin VB.CommandButton cmdSignDistortVerify 
                  Caption         =   "Sign, Distort, and Verify"
                  Height          =   375
                  Left            =   9930
                  TabIndex        =   122
                  Top             =   780
                  Width           =   3045
               End
               Begin ComctlLib.TreeView trvHashResults 
                  Height          =   2265
                  Left            =   60
                  TabIndex        =   121
                  Top             =   690
                  Width           =   5970
                  _ExtentX        =   10530
                  _ExtentY        =   3995
                  _Version        =   327682
                  Style           =   7
                  Appearance      =   1
               End
               Begin VB.TextBox txtPreImage 
                  Height          =   285
                  Left            =   6225
                  TabIndex        =   119
                  Top             =   345
                  Width           =   6780
               End
               Begin VB.ComboBox cboHashAlgs 
                  Height          =   330
                  ItemData        =   "frmMain.frx":5FC3
                  Left            =   60
                  List            =   "frmMain.frx":5FC5
                  Style           =   2  'Dropdown List
                  TabIndex        =   117
                  Top             =   330
                  Width           =   6000
               End
               Begin VB.Label Label41 
                  BackColor       =   &H80000009&
                  Caption         =   $"frmMain.frx":5FC7
                  Height          =   1545
                  Left            =   6135
                  TabIndex        =   125
                  Top             =   1380
                  Width           =   6855
               End
               Begin VB.Label lblMain 
                  BackColor       =   &H80000009&
                  Caption         =   "Pre-image:"
                  Height          =   255
                  Index           =   2
                  Left            =   6225
                  TabIndex        =   120
                  Top             =   75
                  Width           =   1215
               End
               Begin VB.Label lblMain 
                  BackColor       =   &H80000009&
                  Caption         =   "Hash Algorithms:"
                  Height          =   255
                  Index           =   4
                  Left            =   60
                  TabIndex        =   118
                  Top             =   60
                  Width           =   1815
               End
            End
            Begin VB.PictureBox picCrypt 
               BackColor       =   &H80000009&
               BorderStyle     =   0  'None
               Height          =   3015
               Index           =   0
               Left            =   90
               ScaleHeight     =   3015
               ScaleWidth      =   13050
               TabIndex        =   109
               Top             =   405
               Width           =   13050
               Begin ComctlLib.TreeView trvProviders 
                  Height          =   2295
                  Left            =   45
                  TabIndex        =   114
                  Top             =   675
                  Width           =   6120
                  _ExtentX        =   10795
                  _ExtentY        =   4048
                  _Version        =   327682
                  Style           =   7
                  Appearance      =   1
               End
               Begin VB.TextBox txtContainerName 
                  Height          =   285
                  Left            =   6285
                  TabIndex        =   112
                  Top             =   315
                  Width           =   6720
               End
               Begin VB.ComboBox cboProviders 
                  Height          =   330
                  ItemData        =   "frmMain.frx":6119
                  Left            =   45
                  List            =   "frmMain.frx":611B
                  Style           =   2  'Dropdown List
                  TabIndex        =   110
                  Top             =   315
                  Width           =   6150
               End
               Begin VB.Label Label40 
                  BackColor       =   &H80000009&
                  Caption         =   $"frmMain.frx":611D
                  Height          =   2205
                  Left            =   6240
                  TabIndex        =   115
                  Top             =   750
                  Width           =   6750
               End
               Begin VB.Label Label39 
                  BackColor       =   &H80000009&
                  Caption         =   "Container Name:"
                  Height          =   225
                  Left            =   6285
                  TabIndex        =   113
                  Top             =   75
                  Width           =   1725
               End
               Begin VB.Label Label38 
                  BackColor       =   &H80000009&
                  Caption         =   "Providers:"
                  Height          =   225
                  Left            =   45
                  TabIndex        =   111
                  Top             =   75
                  Width           =   1320
               End
            End
            Begin ComctlLib.TabStrip TabCrypt 
               Height          =   3390
               Left            =   30
               TabIndex        =   108
               Top             =   60
               Width           =   13155
               _ExtentX        =   23204
               _ExtentY        =   5980
               _Version        =   327682
               BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
                  NumTabs         =   3
                  BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Providers"
                     Key             =   "Providers"
                     Object.Tag             =   ""
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Hash"
                     Key             =   "Hash"
                     Object.Tag             =   ""
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Keys"
                     Key             =   "Keys"
                     Object.Tag             =   ""
                     ImageVarType    =   2
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3495
            Index           =   3
            Left            =   90
            ScaleHeight     =   3495
            ScaleWidth      =   13215
            TabIndex        =   82
            Top             =   420
            Visible         =   0   'False
            Width           =   13215
            Begin VB.CommandButton cmd_ShowCertificate 
               Caption         =   "Show Cert"
               Height          =   330
               Left            =   6690
               TabIndex        =   107
               Top             =   3015
               Width           =   1470
            End
            Begin VB.CommandButton cmdShowCert 
               Caption         =   "Get{}"
               Height          =   345
               Left            =   7080
               TabIndex        =   103
               Top             =   2295
               Width           =   795
            End
            Begin VB.CommandButton cmdGetSign 
               Caption         =   "Get Sign"
               Height          =   330
               Left            =   4980
               TabIndex        =   102
               Top             =   3015
               Width           =   1470
            End
            Begin VB.CommandButton cmdSingFile 
               Caption         =   "Sign File"
               Height          =   330
               Left            =   3240
               TabIndex        =   101
               Top             =   3015
               Width           =   1470
            End
            Begin VB.TextBox txtDescriptionURL 
               Height          =   270
               Left            =   1800
               TabIndex        =   99
               Top             =   2520
               Width           =   5145
            End
            Begin VB.TextBox txtDescription 
               Height          =   270
               Left            =   1365
               TabIndex        =   98
               Top             =   2115
               Width           =   5595
            End
            Begin VB.CommandButton cmdTimeStamp 
               Caption         =   "Time Stamp"
               Height          =   330
               Left            =   1515
               TabIndex        =   96
               Top             =   3015
               Width           =   1470
            End
            Begin VB.CommandButton cmdVerify 
               Caption         =   "Verify"
               Height          =   330
               Left            =   150
               TabIndex        =   95
               Top             =   3015
               Width           =   1095
            End
            Begin VB.TextBox txtFileToSing 
               Height          =   270
               Left            =   1185
               Locked          =   -1  'True
               TabIndex        =   93
               Top             =   1680
               Width           =   5190
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "..."
               Height          =   240
               Left            =   6465
               TabIndex        =   92
               Top             =   1680
               Width           =   480
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H80000009&
               Caption         =   "Sign Certificate"
               Height          =   1560
               Left            =   60
               TabIndex        =   83
               Top             =   45
               Width           =   6915
               Begin VB.TextBox txtCer 
                  Height          =   285
                  Index           =   3
                  Left            =   2490
                  Locked          =   -1  'True
                  TabIndex        =   91
                  Top             =   1200
                  Width           =   4350
               End
               Begin VB.TextBox txtCer 
                  Height          =   285
                  Index           =   2
                  Left            =   2490
                  Locked          =   -1  'True
                  TabIndex        =   90
                  Top             =   885
                  Width           =   4350
               End
               Begin VB.TextBox txtCer 
                  Height          =   285
                  Index           =   1
                  Left            =   2490
                  Locked          =   -1  'True
                  TabIndex        =   89
                  Top             =   555
                  Width           =   4350
               End
               Begin VB.TextBox txtCer 
                  Height          =   285
                  Index           =   0
                  Left            =   2490
                  Locked          =   -1  'True
                  TabIndex        =   88
                  Top             =   240
                  Width           =   4350
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000009&
                  Caption         =   "Certificate Type:"
                  Height          =   240
                  Left            =   105
                  TabIndex        =   87
                  Top             =   1140
                  Width           =   2310
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000009&
                  Caption         =   "Certificate Location:"
                  Height          =   240
                  Left            =   105
                  TabIndex        =   86
                  Top             =   840
                  Width           =   2310
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000009&
                  Caption         =   "Certificate Store:"
                  Height          =   240
                  Left            =   105
                  TabIndex        =   85
                  Top             =   555
                  Width           =   2310
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000009&
                  Caption         =   "Certificate Name:"
                  Height          =   240
                  Left            =   105
                  TabIndex        =   84
                  Top             =   270
                  Width           =   2310
               End
            End
            Begin VB.Label Label37 
               BackColor       =   &H80000009&
               Caption         =   "Certificate SCHEMA -->"
               Height          =   270
               Left            =   7095
               TabIndex        =   105
               Top             =   1890
               Width           =   2445
            End
            Begin VB.Label Label36 
               BackColor       =   &H80000009&
               Caption         =   $"frmMain.frx":6311
               Height          =   1830
               Left            =   7095
               TabIndex        =   104
               Top             =   165
               Width           =   2385
            End
            Begin VB.Image Image3 
               Height          =   3765
               Left            =   9525
               Picture         =   "frmMain.frx":63AD
               Top             =   -45
               Width           =   3660
            End
            Begin VB.Label Label35 
               BackColor       =   &H80000009&
               Caption         =   "Description URL:"
               Height          =   225
               Left            =   75
               TabIndex        =   100
               Top             =   2550
               Width           =   1755
            End
            Begin VB.Label Label34 
               BackColor       =   &H80000009&
               Caption         =   "Description:"
               Height          =   225
               Left            =   75
               TabIndex        =   97
               Top             =   2115
               Width           =   1305
            End
            Begin VB.Label Label33 
               BackColor       =   &H80000009&
               Caption         =   "File Name:"
               Height          =   225
               Left            =   75
               TabIndex        =   94
               Top             =   1695
               Width           =   1185
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3480
            Index           =   2
            Left            =   90
            ScaleHeight     =   3480
            ScaleWidth      =   13230
            TabIndex        =   66
            Top             =   420
            Visible         =   0   'False
            Width           =   13230
            Begin VB.Frame Frame5 
               BackColor       =   &H80000009&
               Caption         =   "Info Certificate"
               Height          =   3405
               Left            =   7560
               TabIndex        =   79
               Top             =   30
               Width           =   5625
               Begin VB.PictureBox Picture5 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3165
                  Left            =   30
                  ScaleHeight     =   3165
                  ScaleWidth      =   5550
                  TabIndex        =   80
                  Top             =   195
                  Width           =   5550
                  Begin VB.TextBox txtInfCert 
                     BorderStyle     =   0  'None
                     Height          =   3090
                     Left            =   30
                     Locked          =   -1  'True
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   81
                     Top             =   30
                     Width           =   5475
                  End
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H80000009&
               Caption         =   "Certificates Store"
               Height          =   3405
               Left            =   15
               TabIndex        =   67
               Top             =   30
               Width           =   7500
               Begin VB.PictureBox Picture4 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   3165
                  Left            =   45
                  ScaleHeight     =   3165
                  ScaleWidth      =   7410
                  TabIndex        =   68
                  Top             =   195
                  Width           =   7410
                  Begin VB.CommandButton cmdFind 
                     Caption         =   "&Find Certificate"
                     Height          =   330
                     Left            =   5220
                     TabIndex        =   78
                     Top             =   75
                     Width           =   2085
                  End
                  Begin VB.TextBox txtCriteria 
                     Height          =   315
                     Left            =   2115
                     TabIndex        =   76
                     Top             =   1290
                     Width           =   5235
                  End
                  Begin VB.ComboBox cmbFindType 
                     Height          =   330
                     ItemData        =   "frmMain.frx":9380
                     Left            =   2100
                     List            =   "frmMain.frx":93A8
                     Style           =   2  'Dropdown List
                     TabIndex        =   74
                     Top             =   900
                     Width           =   5265
                  End
                  Begin VB.ComboBox cmbStoreLocation 
                     Height          =   330
                     ItemData        =   "frmMain.frx":958C
                     Left            =   2100
                     List            =   "frmMain.frx":959C
                     Style           =   2  'Dropdown List
                     TabIndex        =   72
                     Top             =   510
                     Width           =   5265
                  End
                  Begin ComctlLib.ListView lstFoundCerts 
                     Height          =   1410
                     Left            =   15
                     TabIndex        =   71
                     Top             =   1725
                     Width           =   7350
                     _ExtentX        =   12965
                     _ExtentY        =   2487
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   327682
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   4
                     BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                        Key             =   ""
                        Object.Tag             =   ""
                        Text            =   "Subject"
                        Object.Width           =   11995
                     EndProperty
                     BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                        SubItemIndex    =   1
                        Key             =   ""
                        Object.Tag             =   ""
                        Text            =   "Issuer"
                        Object.Width           =   2540
                     EndProperty
                     BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                        SubItemIndex    =   2
                        Key             =   ""
                        Object.Tag             =   ""
                        Text            =   "Valid From"
                        Object.Width           =   2540
                     EndProperty
                     BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                        SubItemIndex    =   3
                        Key             =   ""
                        Object.Tag             =   ""
                        Text            =   "Valid To"
                        Object.Width           =   2540
                     EndProperty
                  End
                  Begin VB.ComboBox cmbStoreName 
                     Height          =   330
                     ItemData        =   "frmMain.frx":961D
                     Left            =   2100
                     List            =   "frmMain.frx":962D
                     Style           =   2  'Dropdown List
                     TabIndex        =   69
                     Top             =   150
                     Width           =   2895
                  End
                  Begin VB.Image Image1 
                     Height          =   480
                     Left            =   150
                     Picture         =   "frmMain.frx":9650
                     Top             =   855
                     Width           =   480
                  End
                  Begin VB.Label Label28 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Search Criteria:"
                     Height          =   255
                     Left            =   180
                     TabIndex        =   77
                     Top             =   1335
                     Width           =   1800
                  End
                  Begin VB.Label Label27 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Type:"
                     Height          =   210
                     Left            =   75
                     TabIndex        =   75
                     Top             =   960
                     Width           =   1905
                  End
                  Begin VB.Label Label26 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000009&
                     Caption         =   "Store Location:"
                     Height          =   225
                     Left            =   75
                     TabIndex        =   73
                     Top             =   600
                     Width           =   1905
                  End
                  Begin VB.Label Label25 
                     BackColor       =   &H80000009&
                     Caption         =   "Certificate Store:"
                     Height          =   255
                     Left            =   75
                     TabIndex        =   70
                     Top             =   180
                     Width           =   1965
                  End
               End
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3480
            Index           =   1
            Left            =   90
            ScaleHeight     =   3480
            ScaleWidth      =   13230
            TabIndex        =   39
            Top             =   420
            Visible         =   0   'False
            Width           =   13230
            Begin ProjectCapicom.Calendario Calendario 
               Height          =   240
               Left            =   75
               TabIndex        =   65
               Top             =   120
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   423
               LongData        =   "giovedì 2 aprile 2009"
               ShortData       =   "02/04/2009"
            End
            Begin VB.TextBox txtKeyPassWord 
               Height          =   315
               Left            =   11115
               TabIndex        =   64
               Top             =   3105
               Width           =   2010
            End
            Begin VB.CommandButton cmdDecryptKey 
               Caption         =   "Decrypt Key"
               Height          =   375
               Left            =   8370
               TabIndex        =   62
               ToolTipText     =   "Decrypt from file key"
               Top             =   3030
               Width           =   1485
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000009&
               Caption         =   "Data Crypted/Encrypted"
               Height          =   2880
               Left            =   6405
               TabIndex        =   59
               Top             =   30
               Width           =   6780
               Begin VB.PictureBox Picture3 
                  BackColor       =   &H80000009&
                  BorderStyle     =   0  'None
                  Height          =   2625
                  Left            =   45
                  ScaleHeight     =   2625
                  ScaleWidth      =   6690
                  TabIndex        =   60
                  Top             =   210
                  Width           =   6690
                  Begin VB.TextBox txtDataKey 
                     Height          =   2535
                     Left            =   15
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   61
                     Top             =   45
                     Width           =   6630
                  End
               End
            End
            Begin VB.CommandButton cmdCreateKey 
               Caption         =   "Create Key"
               Height          =   375
               Left            =   6720
               TabIndex        =   58
               ToolTipText     =   "Create file Key Crypted"
               Top             =   3030
               Width           =   1425
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Index           =   7
               Left            =   1710
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   56
               Top             =   2835
               Width           =   4635
            End
            Begin VB.CommandButton cmdGenKey 
               Caption         =   "..."
               Height          =   285
               Left            =   5790
               TabIndex        =   55
               ToolTipText     =   "Generate Seial Key"
               Top             =   2070
               Width           =   525
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1710
               TabIndex        =   47
               Top             =   90
               Width           =   4620
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   315
               Index           =   1
               Left            =   1710
               TabIndex        =   46
               Top             =   495
               Width           =   4620
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   2
               Left            =   1710
               TabIndex        =   45
               Top             =   2445
               Width           =   4185
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1710
               TabIndex        =   44
               Top             =   1305
               Width           =   4620
            End
            Begin VB.TextBox ts 
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
               Index           =   4
               Left            =   1710
               TabIndex        =   43
               Top             =   1680
               Width           =   4620
            End
            Begin VB.TextBox ts 
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
               Index           =   5
               Left            =   1710
               TabIndex        =   42
               Top             =   2055
               Width           =   4050
            End
            Begin VB.TextBox ts 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   1710
               TabIndex        =   41
               Top             =   900
               Width           =   4605
            End
            Begin VB.Label Label24 
               BackColor       =   &H80000009&
               Caption         =   "PassWord:"
               Height          =   225
               Left            =   10065
               TabIndex        =   63
               Top             =   3135
               Width           =   1020
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Note:"
               Height          =   225
               Left            =   60
               TabIndex        =   57
               Top             =   2850
               Width           =   1515
            End
            Begin VB.Image imgCF 
               Height          =   240
               Left            =   6030
               Picture         =   "frmMain.frx":9F1A
               ToolTipText     =   "Show Codice Fiscale (Only for Italy) sorry!"
               Top             =   2475
               Width           =   240
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Validity Date:"
               Height          =   225
               Left            =   60
               TabIndex        =   54
               Top             =   975
               Width           =   1515
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Signatory:"
               Height          =   225
               Left            =   60
               TabIndex        =   53
               Top             =   120
               Width           =   1515
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Certifier:"
               Height          =   225
               Left            =   60
               TabIndex        =   52
               Top             =   525
               Width           =   1515
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CF:"
               Height          =   225
               Left            =   60
               TabIndex        =   51
               Top             =   2490
               Width           =   1515
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               Height          =   225
               Left            =   60
               TabIndex        =   50
               Top             =   1335
               Width           =   1515
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "State:"
               Height          =   225
               Left            =   60
               TabIndex        =   49
               Top             =   1725
               Width           =   1515
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Signature ID:"
               Height          =   225
               Left            =   60
               TabIndex        =   48
               Top             =   2100
               Width           =   1515
            End
         End
         Begin VB.PictureBox picsTabs 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   3465
            Index           =   0
            Left            =   90
            ScaleHeight     =   3465
            ScaleWidth      =   13215
            TabIndex        =   11
            Top             =   450
            Width           =   13215
            Begin VB.CommandButton cmdDecode 
               Caption         =   "Decode"
               Height          =   375
               Left            =   6075
               TabIndex        =   40
               ToolTipText     =   "Open and Decrypt File"
               Top             =   2175
               Width           =   1050
            End
            Begin VB.CommandButton cmdEncodeF 
               Caption         =   "Encode"
               Height          =   375
               Left            =   6075
               TabIndex        =   38
               ToolTipText     =   "Open and Crypt File"
               Top             =   1545
               Width           =   1050
            End
            Begin VB.CheckBox CheckShowPsW 
               BackColor       =   &H80000009&
               Height          =   225
               Left            =   5430
               TabIndex        =   33
               ToolTipText     =   "Show/Hide PassWord"
               Top             =   3150
               Width           =   210
            End
            Begin VB.CommandButton cmdRandom 
               Caption         =   "..."
               Height          =   255
               Left            =   5805
               TabIndex        =   32
               ToolTipText     =   "Generate Random PassWord"
               Top             =   3135
               Width           =   450
            End
            Begin VB.CheckBox CheckTags 
               BackColor       =   &H80000009&
               Caption         =   "Include {Begin-End} Tags"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2880
               TabIndex        =   21
               Top             =   2640
               Width           =   3045
            End
            Begin VB.TextBox txtFileName 
               Height          =   285
               Left            =   7575
               TabIndex        =   20
               Top             =   3150
               Width           =   5595
            End
            Begin VB.CommandButton cmdDecrypt 
               Caption         =   "Dencrypt"
               Height          =   375
               Left            =   6075
               TabIndex        =   18
               Top             =   915
               Width           =   1050
            End
            Begin VB.CommandButton cmdEncrypt 
               Caption         =   "Encrypt"
               Height          =   375
               Left            =   6075
               TabIndex        =   17
               Top             =   270
               Width           =   1050
            End
            Begin VB.CheckBox CheckToFile 
               BackColor       =   &H80000009&
               Caption         =   "Encrypt and create File"
               Height          =   255
               Left            =   30
               TabIndex        =   16
               Top             =   2640
               Width           =   2820
            End
            Begin VB.TextBox txtPassword 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1095
               TabIndex        =   15
               Text            =   "{5B439EA9-5E69-402D-89CD-4659BA763866}"
               Top             =   3135
               Width           =   4230
            End
            Begin VB.TextBox txtOut 
               Height          =   2550
               Left            =   7260
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   45
               Width           =   5910
            End
            Begin VB.TextBox txtIn 
               Height          =   2550
               Left            =   45
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Text            =   "frmMain.frx":A2A4
               Top             =   45
               Width           =   5910
            End
            Begin VB.Label lblCRC 
               BackColor       =   &H80000009&
               Height          =   270
               Left            =   7950
               TabIndex        =   37
               Top             =   2865
               Width           =   5220
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000009&
               Caption         =   "CRC32:"
               Height          =   255
               Left            =   7245
               TabIndex        =   36
               Top             =   2865
               Width           =   645
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000009&
               Caption         =   "HASH:"
               Height          =   255
               Left            =   7275
               TabIndex        =   35
               Top             =   2610
               Width           =   585
            End
            Begin VB.Label lblHash 
               BackColor       =   &H80000009&
               Height          =   270
               Left            =   7950
               TabIndex        =   34
               Top             =   2610
               Width           =   5220
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000009&
               Caption         =   "FileName:"
               Height          =   255
               Left            =   6495
               TabIndex        =   19
               Top             =   3180
               Width           =   975
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000009&
               Caption         =   "PassWord:"
               Height          =   255
               Left            =   90
               TabIndex        =   14
               Top             =   3135
               Width           =   1020
            End
         End
         Begin ComctlLib.TabStrip TBS 
            Height          =   3870
            Left            =   60
            TabIndex        =   10
            Top             =   90
            Width           =   13290
            _ExtentX        =   23442
            _ExtentY        =   6826
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   7
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Encrypt/Decrypt"
                  Key             =   "EncDec"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Serial Key type Encryption"
                  Key             =   "SerialKey"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Certificate"
                  Key             =   "Certificate"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Sign your Software"
                  Key             =   "SignSoftware"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "CryptoAPI Function"
                  Key             =   "CryptoAPI"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Create Certificate"
                  Key             =   "CreateCertificate"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Sign MIME"
                  Key             =   "SignMIME"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CryptoAPI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   480
      Left            =   4995
      TabIndex        =   153
      Top             =   -15
      Width           =   1935
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CryptoAPI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   4290
      TabIndex        =   152
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":A444
      Height          =   495
      Left            =   90
      TabIndex        =   7
      Top             =   1050
      Width           =   13365
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   195
      Left            =   2775
      TabIndex        =   6
      Top             =   825
      Width           =   7755
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CAPICOM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   405
      Left            =   915
      TabIndex        =   5
      Top             =   510
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CAPICOM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   405
      Left            =   135
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   9495
      Picture         =   "frmMain.frx":A515
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3120
      TabIndex        =   3
      Top             =   195
      Width           =   1005
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   300
      Left            =   3150
      TabIndex        =   2
      Top             =   210
      Width           =   1005
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Signature Certified"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   210
      TabIndex        =   1
      Top             =   45
      Width           =   3210
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Signature Certified"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   75
      Width           =   3210
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   -765
      Picture         =   "frmMain.frx":ADDF
      Top             =   -150
      Width           =   14325
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   13575
   End
   Begin VB.Menu mnuCertificate 
      Caption         =   "Certificate"
      Visible         =   0   'False
      Begin VB.Menu mnuCert 
         Caption         =   "Show"
         Index           =   0
      End
      Begin VB.Menu mnuCert 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu mnuCert 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCert 
         Caption         =   "Use this"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*/ .... About CAPICOM:
' .... One of the primary objectives during the development of CAPICOM was to make it as easy as possible certain
' .... cryptographic operations. To affix a digital signature, leaving out the necessary checks on errors, would be
' .... sufficient only three lines:
'
' .... Dim signed As New SignedData
' .... signed.Content = bufferToSign
' .... MsgBox signed.Sign…
'\*

Option Explicit

' .... Time Stamping URL
Private Const URL = "http://timestamp.verisign.com/scripts/timstamp.dll"
Private Signer As New Signer
Private SignedCode As New SignedCode
    
Private Const tTitle As String = "Signature Certified v1.0.2ß"

Private tbsKey As String
Private cfPath As String

' .... Class GUIDE
Dim objGUIDE As clsGUIDGenerator

' .... Disabled Closed
Dim readyToClose As Boolean

Dim ssLeft As Long
Dim ssTop As Long

' .... CryptoAPI
Private mbytHashValue() As Byte
Private mlngHProvider As Long
Private mlngHHash As Long
Private mlngHKey As Long

' .... MIME Message
Private Const CERT_KEY_SPEC_PROP_ID = 6
Private Const CdoAddressListGAL = 0
Private Const CdoAddressListPAB = 1
Private Const cdlOFNFileMustExist = &H1000
Private Const cdlOFNHideReadOnly = &H4
Private Const cdlOFNPathMustExist = &H800
Private Const cdlOFNCreatePrompt = &H2000

Private oSigner As New CAPICOM.Signer
Private oMessage As New CDO.Message
Private Sub Calendario_DblClick()
    ts(6).Text = Format(Calendario.ShortData, "mm/dd/yyyy") & " " & Time & " GMT"
End Sub


Private Sub CalVal_DblClick()
    Dim tpiu As Long
    Dim d2 As Long
    ' .... Retrive the Start Date
    If Format(Now, "mm") < Format(CalVal.ShortData, "mm") Or Format(Now, "yyyy") < Format(CalVal.ShortData, "yyyy") Then
            MsgBox "Select the correct Date please!", vbExclamation, App.Title
        Exit Sub
    End If
    d2 = Format(Mid$(CalVal.ShortData, 7, 8), "00")
    tpiu = Mid$(CalVal.ShortData, 9, 10) ' Truncate the last 2 chars of the Date...
    tpiu = tpiu + 10 ' ... and increase the Year by 10 years, format standard certificate
    
    txtValidFrom.Text = Format(CalVal.ShortData, "mm/dd/yyyy")
    txtValidTo.Text = Format(CalVal.ShortData, "mm/dd/") & Mid$(d2, 1, 2) & tpiu
    
    txtValidFrom.ToolTipText = Format(txtValidFrom.Text, "Long Date")
    txtValidTo.ToolTipText = Format(txtValidTo.Text, "Long Date")
End Sub


Private Sub cboProviders_Click()
    GetProviderInfo
End Sub

Private Sub CheckShowPsW_Click()
    If CheckShowPsW.value = 1 Then txtPassword.PasswordChar = "*" Else txtPassword.PasswordChar = ""
End Sub

Private Sub CheckToFile_Click()
    If CheckToFile.value = 1 Then _
    CheckTags.Enabled = True Else CheckTags.Enabled = False
End Sub

Private Sub cmbAlgorithm_Click()
    If txtOut.Text <> Empty Then
        Dim objCap As New clsCapicom
        lblHash.Caption = objCap.GetHash(txtOut.Text, cmbAlgorithm.ItemData(cmbAlgorithm.ListIndex))
        Set objCap = Nothing
    End If
End Sub


Private Sub cmd_ShowCertificate_Click()
Dim cert As Certificate
    Dim certs As New Certificates
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim StoreName As String
    Dim st As New Store
    On Local Error GoTo ErrorShowCert
        Select Case txtCer(2).Text
         Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
           StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
         Case "CAPICOM_CURRENT_USER_STORE"
            StoreLocation = CAPICOM_CURRENT_USER_STORE
         Case "CAPICOM_LOCAL_MACHINE_STORE"
            StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
         Case "CAPICOM_SMART_CARD_USER_STORE"
            StoreLocation = CAPICOM_SMART_CARD_USER_STORE
         Case Else
            Exit Sub
        End Select
        StoreName = txtCer(1).Text
        st.Open StoreLocation, StoreName, CAPICOM_STORE_OPEN_READ_ONLY
        Set certs = st.Certificates
        If certs.count = 0 Then Exit Sub
        For Each cert In certs
            If LCase(txtCer(0).Text) = LCase(cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME)) Then
                    cert.Display
                Exit For
            End If
        Next cert
    Set certs = Nothing
Exit Sub
ErrorShowCert:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbInformation, App.Title
    Err.Clear
End Sub

Private Sub cmdAdd_Click()
    On Local Error GoTo ErrorCanc
    With cDialog
        .CancelError = True
        .DialogTitle = "Select the File to be Added:"
        .Filter = "All Files (*.*)|*.*"
        .DefaultExt = "*.*"
        .ShowOpen
        If .FileName = Empty Then Exit Sub
        lstList.AddItem .FileName
    End With
Exit Sub
ErrorCanc:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdCreate_Click()
    CreateKey
End Sub

Private Sub cmdCreateCab_Click()
    Dim makeCabPath As String
    On Local Error GoTo ErrorMakeCab
    makeCabPath = App.Path + "\tools\Makecab.exe"
    If Dir$(makeCabPath) = Empty Then
            MsgBox "File {Makecab.exe} is missing!", vbExclamation, App.Title
        Exit Sub
    End If
    If Dir$(txtDestPath.Text + "_setup.xml") = Empty Then
            MsgBox "Create first the _setup file (*.xml)!", vbExclamation, App.Title
        Exit Sub
    End If
    If lstList.ListCount = 0 Then lstList.AddItem txtDestPath.Text + "_setup.xml"
    If picCab.Visible = False Then picCab.Visible = True Else picCab.Visible = False
Exit Sub
ErrorMakeCab:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

' .... Now create a Trusted Certificate and apply this on your HTTPs server
Private Sub cmdCreateCertificate_Click()
    Dim intFile As Integer
    Dim i As Integer
    Dim sStore As String
    Dim makecertpath As String
    Dim pvk2pfxpath As String
    Dim pvkimprtpath As String
    Dim openSSLpath As String
    Dim strAutority As String
    Dim strAutorityReleased As String
    Dim strUse As String
    On Local Error GoTo ErrorCreate
    
    makecertpath = App.Path + "\tools\makecert.exe"
    pvk2pfxpath = App.Path + "\tools\pvk2pfx.exe"
    pvkimprtpath = App.Path + "\tools\pvkimprt.exe"
    openSSLpath = App.Path + "\tools\openssl.exe"
    
    sStore = cmb_Stores.List(cmb_Stores.ListIndex)
    
    If sStore = Empty Then
            MsgBox "Please, select the Store Certificate!", vbExclamation, App.Title
        Exit Sub
    End If
    
    If Dir$(makecertpath) = Empty Then
            MsgBox "File {makecert.exe} is missing!", vbExclamation, App.Title
        Exit Sub
    ElseIf Dir$(pvk2pfxpath) = Empty Then
            MsgBox "File {pvk2pfx.exe} is missing!", vbExclamation, App.Title
        Exit Sub
    ElseIf Dir$(pvkimprtpath) = Empty Then
            MsgBox "File {pvkimprtpath} is missing!", vbExclamation, App.Title
        Exit Sub
    ElseIf Dir$(App.Path + "\tools\signer.dll") = Empty Then
            MsgBox "Library {signer.dll} is missing!", vbExclamation, App.Title
        Exit Sub
    End If
    If txtCertificateName.Text = Empty Then
            MsgBox "Certificate Name required!", vbExclamation, App.Title
        Exit Sub
    End If
    If txtAutority.Text = Empty Then
            MsgBox "Certificate Autority required!", vbExclamation, App.Title
        Exit Sub
    End If
    
    If Dir$(txtDestPath.Text) = Empty Then
        If MakeDirectory(txtDestPath.Text) = False Then
                MsgBox "Select the Default path of Certificates!", vbExclamation, App.Title
            Exit Sub
        End If
    End If
    
    strUse = Empty
    For i = 0 To 6
        If CheckUsers(i).value = 1 Then strUse = strUse & CheckUsers(i).Caption & ","
    Next i
    
    If strUse <> Empty And Mid$(strUse, 1, 1) <> "," Then
        ' .... Remuve last ","
        strUse = Mid$(strUse, 1, Len(strUse) - 1)
    Else
    ' .... If not Checked use Default = Code Signing and Time Stamping
        strUse = "1.3.6.1.5.5.7.3.3,1.3.6.1.5.5.7.3.8"
    End If
    
    ' .... In the case of a PC with dynamic IP, you must indicate, instead of 'www.yourserver.com'
    ' .... the name of the computer, for example 'mypc'
    If CheckIP.value = 1 Then _
        strAutority = txtInternalIPs.List(txtInternalIPs.ListIndex) Else _
        strAutority = txtAutority.Text
    
    strAutorityReleased = txtAutority.Text
    
    ' .... Create batch file
    intFile = FreeFile()
    Open App.Path + "\batch.bat" For Output As #intFile
        ' .... Standard Validity Ex: Now -> to 2040
        ' .... if Checked include your Validity ;) Ex: Now -> to 2019
        If CheckPutMyValidity.value = 0 Then _
        Print #intFile, """"; makecertpath; """" & " -pe -n " & """CN=" & strAutority & """" & " -r -a sha1 -sky signature -sv " _
        & """"; txtDestPath.Text & txtCertificateName.Text & ".pvk"""; " """; txtDestPath.Text & txtCertificateName.Text & ".cer"""; _
        Else Print #intFile, """"; makecertpath; """" & " -pe -n " & """CN=" & strAutority & """" & " -r -a sha1 -sky signature -b " _
        & txtValidFrom.Text & " -e " & txtValidTo.Text & " -sv " & """"; txtDestPath.Text & txtCertificateName.Text _
        & ".pvk"""; " """; txtDestPath.Text & txtCertificateName.Text & ".cer""";
    Close #intFile
    
    If Dir$(App.Path + "\batch.bat") = Empty Then
            MsgBox "Error to create batch procedure!", vbExclamation, App.Title
        Exit Sub
    End If
    ' .... Request create Certificates
    If MsgBox("Now create Certificate. Will be prompted for a password! Continue?", vbYesNo + vbInformation + _
        vbDefaultButton1, App.Title) = vbNo Then
    ' .... cancel? Delete batch file
            If Dir$(App.Path + "\batch.bat") <> Empty Then Call Kill(App.Path + "\batch.bat")
        Exit Sub
    End If
    
    txtBase64.Text = Empty
    txtSHA1.Text = Empty
    
    ' .... Run Step 1
    If ShelledAPP(App.Path + "\batch.bat", vbHide) = "End" Then
        MsgBox "Step-1) Batch process terminate success! Ok to Continue!", vbInformation, App.Title
    Else
        MsgBox "Batch process terminate with Error!", vbExclamation, App.Title
    ' .... Error? Delete batch file
            If Dir$(App.Path + "\batch.bat") <> Empty Then Call Kill(App.Path + "\batch.bat")
        Exit Sub
    End If
    
    Open App.Path + "\batch.bat" For Output As #intFile
        If CheckInstall.value = 1 Then
            Print #intFile, """"; makecertpath; """" & " -pe -ic " & """"; txtDestPath.Text & txtCertificateName.Text & ".cer"""; " -iv " & """"; txtDestPath.Text & txtCertificateName.Text & ".pvk"""; " -iky signature -sky exchange -b " & txtValidFrom.Text & " -e " & txtValidTo.Text & " -eku " & strUse & " -ss " & sStore & " -sp "; """Microsoft RSA SChannel Cryptographic Provider"""; " -n "; """CN=SVCWCFCER"""; " -sy 12 -sv" & " """; txtDestPath.Text & txtCertificateName.Text & "_.pvk"""; " """; txtDestPath.Text & txtCertificateName.Text & "_.cer""";
        Else
            Print #intFile, """"; makecertpath; """" & " -pe -ic " & """"; txtDestPath.Text & txtCertificateName.Text & ".cer"""; " -iv " & """"; txtDestPath.Text & txtCertificateName.Text & ".pvk"""; " -iky signature -sky exchange -b " & txtValidFrom.Text & " -e " & txtValidTo.Text & " -eku " & strUse & " -sp "; """Microsoft RSA SChannel Cryptographic Provider"""; " -n "; """CN=SVCWCFCER"""; " -sy 12 -sv" & " """; txtDestPath.Text & txtCertificateName.Text & "_.pvk"""; " """; txtDestPath.Text & txtCertificateName.Text & "_.cer""";
        End If
        ' .... SVCWCFCER -> to your Name not create Certificate
        'Print #intFile, """"; makecertpath; """" & " -pe -ic " & """"; txtDestPath.Text & txtCertificateName.Text & ".cer"""; " -iv " & """"; txtDestPath.Text & txtCertificateName.Text & ".pvk"""; " -iky signature -sky exchange -b " & txtValidFrom.Text & " -e " & txtValidTo.Text & " -eku " & strUse & " -sp "; """Microsoft RSA SChannel Cryptographic Provider"""; " -n "; """CN=" & strAutorityReleased & """"; " - sy 12 - sv " & " """; txtDestPath.Text & txtCertificateName.Text & "_.pvk"""; " """; txtDestPath.Text & txtCertificateName.Text & "_.cer""";
    Close #intFile
    
    ' .... Run Step 2
    If ShelledAPP(App.Path + "\batch.bat", vbHide) = "End" Then
        MsgBox "Step-2) Batch process terminate success! Ok to Continue!", vbInformation, App.Title
    Else
        MsgBox "Batch process terminate with Error!", vbExclamation, App.Title
    ' .... Error? Delete batch file
            If Dir$(App.Path + "\batch.bat") <> Empty Then Call Kill(App.Path + "\batch.bat")
        Exit Sub
    End If
    
    ' .... Also create *.pfx file?
    If CheckPfx.value = 1 Then
        Open App.Path + "\batch.bat" For Output As #intFile
            Print #intFile, """"; pvk2pfxpath; """" & " -pvk " & """"; txtDestPath.Text & txtCertificateName.Text & "_.pvk"""; " -pi " & txtPassCert.Text & " -spc " & """"; txtDestPath.Text & txtCertificateName.Text & "_.cer"""; " -pfx " & """"; txtDestPath.Text & txtCertificateName.Text & "_.pfx"""; " -po " & txtPassCert.Text
        Close #intFile
    If ShelledAPP(App.Path + "\batch.bat", vbHide) = "End" Then
        MsgBox "Step-3) Batch process terminate. PFX created success! All Certificates created success!" _
        & vbCr & vbCr & "Then you import the New Certificate into Trusted Certificates, so that it is Applicable.", vbInformation, App.Title
    Else
            MsgBox "Batch process terminate with Error!", vbExclamation, App.Title
     ' .... Error? Delete batch file
            If Dir$(App.Path + "\batch.bat") <> Empty Then Call Kill(App.Path + "\batch.bat")
        Exit Sub
    End If
End If
    
    If Dir$(openSSLpath) = Empty Then
            MsgBox "File {openssl.exe} is missing!", vbExclamation, App.Title
    Else
        ' .... Create output SHA1 and Base64 *.txt
        Open App.Path + "\batch.bat" For Output As #intFile
            Print #intFile, """"; openSSLpath; """" & " sha1 " & """"; txtDestPath.Text & txtCertificateName.Text & ".cer"" > "; """"; txtDestPath.Text & txtCertificateName.Text & "_Sha1.txt""";
        Close #intFile
        If ShelledAPP(App.Path + "\batch.bat", vbHide) = "End" Then
            Open App.Path + "\batch.bat" For Output As #intFile
                Print #intFile, """"; openSSLpath; """" & " base64 -in " & """"; txtDestPath.Text & txtCertificateName.Text & ".cer"" > "; """"; txtDestPath.Text & txtCertificateName.Text & "_Base64.txt""";
            Close #intFile
            ' .... Run Silent
            If ShelledAPP(App.Path + "\batch.bat", vbHide) = "End" Then:
        End If
    End If
    
    ' .... Open and Display the Base64 outPut
    If Dir$(txtDestPath.Text & txtCertificateName.Text & "_Base64.txt") <> Empty Then
        Open txtDestPath.Text & txtCertificateName.Text & "_Base64.txt" For Input As #intFile
            txtBase64.Text = Input(LOF(intFile), intFile)
        Close #intFile
    End If
    
    ' .... Open and Display the SHA1 outPut
    If Dir$(txtDestPath.Text & txtCertificateName.Text & "_Sha1.txt") <> Empty Then
        Open txtDestPath.Text & txtCertificateName.Text & "_Sha1.txt" For Input As #intFile
            txtSHA1.Text = StripLeft(Input(LOF(intFile), intFile), "=", False)
        Close #intFile
    End If
    
    ' .... Create file *.xml data Setup
    If txtSHA1.Text <> Empty Then
        If txtBase64.Text <> Empty Then
            CreateXML txtDestPath.Text + "_setup.xml"
            If Dir$(txtDestPath.Text + "_setup.xml") <> Empty Then
                lstList.AddItem txtDestPath.Text + "_setup.xml"
            End If
        End If
    End If
    
    ' .... Delete batch file
    If Dir$(App.Path + "\batch.bat") <> Empty Then Call Kill(App.Path + "\batch.bat")
    
    ' .... Open Folder Certificates
    ShellExecute 0&, vbNullString, txtDestPath.Text, vbNullString, "C:\", sShell.vbNormalFocus
Exit Sub
ErrorCreate:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub
Private Sub cmdCreateHash_Click()
    CreateHash
End Sub

Private Sub cmdCreateKey_Click()
    On Local Error GoTo ErrorKey
    If txtKeyPassWord.Text = Empty Or Len(txtKeyPassWord.Text) < 8 Then
            MsgBox "Password len > 7 chars!", vbExclamation, App.Title
        Exit Sub
    End If
    With cDialog
        .CancelError = True
        .DialogTitle = "Save Encrypted Licence as:"
        .Filter = "Licence File key (*.key)|*.key"
        .FilterIndex = 1
        .ShowSave
        If .FileName = Empty Then Exit Sub
            If SaveToFileEncrypted(.FileName, txtKeyPassWord.Text, "|") Then
                MsgBox "The licence file created success!", vbInformation, App.Title
            Else
                MsgBox "Error to create the licence file!", vbExclamation, App.Title
            End If
    End With
Exit Sub
ErrorKey:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdDecode_Click()
Dim sTemp As String
    Dim intFile As Integer
    Dim sPasw As String
    On Local Error GoTo ErrorCanc
    intFile = FreeFile()
    With cDialog
        .CancelError = True
        .DialogTitle = "Select the File to be Decrypted:"
        .Filter = "All Files (*.*)|*.*"
        .DefaultExt = "*.*"
        .ShowOpen
        If .FileName = Empty Then Exit Sub
        Open .FileName For Binary Access Read Lock Read Write As #intFile
            sTemp = Input(LOF(intFile), intFile)
        Close #intFile
    End With
        ' .... Password?
        sPasw = InputBox("Enter the Decrypted Key:", App.Title, "")
        If sPasw = Empty Then Exit Sub
        ' ....Decrypt and Save in oter File
        With cDialog
            .CancelError = True
            .DialogTitle = "Save Encrypted File as:"
            .Filter = "All Files (*.*)|*.*"
            .FilterIndex = 1
            .ShowSave
        If .FileName = Empty Then Exit Sub
        Open .FileName For Output As #intFile
            Print #intFile, DecryptFromString(sTemp, sPasw)
        Close #intFile
    End With
    MsgBox "File Dencrypted success!", vbInformation, App.Title
Exit Sub
ErrorCanc:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub
Private Sub cmdDecrypt_Click()
    On Local Error GoTo ErrorCanc
    If txtOut.Text = Empty Then Exit Sub
    If CheckToFile.value = 0 Then
        txtIn.Text = DecryptFromString(txtOut.Text, txtPassword.Text)
        txtOut.Text = Empty
    Else
        If txtFileName.Text <> Empty And Dir$(txtFileName.Text) <> Empty Then
            txtIn.Text = DecryptFileFromFile(txtFileName.Text, txtPassword.Text)
            txtOut.Text = Empty
        Else
            With cDialog
            .CancelError = True
            .DialogTitle = "Open Encrypted File:"
            .Filter = "NotePad File (*.txt)|*.txt|Word File (*.doc)|*.doc|Rich Text (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .DefaultExt = ".txt"
            .ShowOpen
            If .FileName = Empty Then Exit Sub
            txtFileName.Text = .FileName
            txtIn.Text = DecryptFileFromFile(txtFileName.Text, txtPassword.Text)
            txtOut.Text = Empty
        End With
        End If
    End If
    lblHash.Caption = Empty
    lblCRC.Caption = Empty
Exit Sub
ErrorCanc:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdDecryptFile_Click()
    CipherFile Decrypt
End Sub

Private Sub cmdDecryptFileEasy_Click()
    CipherFileEasy Decrypt
End Sub

Private Sub cmdDecryptKey_Click()
    On Local Error GoTo ErrorKey
    With cDialog
        .CancelError = True
        .DialogTitle = "Select the licence File:"
        .Filter = "Licence File key (*.key)|*.key"
        .DefaultExt = "*.key"
        .ShowOpen
        If .FileName = Empty Then Exit Sub
        txtDataKey.Text = ReadEncrypedFile(.FileName, txtKeyPassWord.Text, "|")
    End With
    If txtDataKey.Text = Empty Then
        MsgBox "Verify your Password please!", vbExclamation, App.Title
    End If
Exit Sub
ErrorKey:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdEncodeF_Click()
    Dim sTemp As String
    Dim intFile As Integer
    Dim sPasw As String
    On Local Error GoTo ErrorCanc
    intFile = FreeFile()
    With cDialog
        .CancelError = True
        .DialogTitle = "Select the File to be Encoded:"
        .Filter = "All Files (*.*)|*.*"
        .DefaultExt = "*.*"
        .ShowOpen
        If .FileName = Empty Then Exit Sub
        Open .FileName For Binary Access Read Lock Read Write As #intFile
            sTemp = Input(LOF(intFile), intFile)
        Close #intFile
    End With
    ' .... Password?
        sPasw = InputBox("Enter the Encrypted Key:", App.Title, "")
        If sPasw = Empty Then Exit Sub
    ' .... Encode
    sTemp = EncryptString(sPasw, sTemp, CAPICOM_ENCRYPTION_ALGORITHM_AES, _
    CAPICOM_ENCRYPTION_KEY_LENGTH_128_BITS, CAPICOM_ENCODE_BASE64)
    ' .... Save in oter File
    With cDialog
        .CancelError = True
        .DialogTitle = "Save Encrypted File as:"
        .Filter = "All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowSave
        If .FileName = Empty Then Exit Sub
        Open .FileName For Output As #intFile
            Print #intFile, sTemp
        Close #intFile
    End With
    MsgBox "File Encrypted success!", vbInformation, App.Title
Exit Sub
ErrorCanc:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub
Private Sub cmdEncrypt_Click()
    Dim objCap As New clsCapicom
    If txtIn.Text = Empty Then Exit Sub
    On Local Error GoTo ErrorCanc
    If CheckToFile.value = 0 Then
        txtOut.Text = EncryptString(txtPassword.Text, txtIn.Text, cmbEncryption.ItemData(cmbEncryption.ListIndex), _
                                    cmbLength.ItemData(cmbLength.ListIndex), cmbBase.ItemData(cmbBase.ListIndex))
        txtIn.Text = Empty
    Else
        If txtFileName.Text <> Empty And Dir$(txtFileName.Text) <> Empty Then
            If MsgBox("Encrypt current FileName (Re-write)?", vbYesNo + vbInformation + _
                vbDefaultButton2, App.Title) = vbYes Then
                txtOut.Text = EncryptFile(txtFileName.Text, txtPassword.Text, txtIn.Text, CBool(CheckTags.value), _
                    cmbEncryption.ItemData(cmbEncryption.ListIndex), cmbLength.ItemData(cmbLength.ListIndex), _
                        cmbBase.ItemData(cmbBase.ListIndex))
                    lblCRC.Caption = GetCRC32(txtFileName.Text)
                    lblHash.Caption = objCap.GetHash(txtOut.Text, cmbAlgorithm.ItemData(cmbAlgorithm.ListIndex))
                    Set objCap = Nothing
                    txtIn.Text = Empty
                Exit Sub
           End If
        End If
        With cDialog
            .CancelError = True
            .DialogTitle = "Save Encrypted File as:"
            .Filter = "NotePad File (*.txt)|*.txt|Word File (*.doc)|*.doc|Rich Text (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .FilterIndex = 1
            .ShowSave
            If .FileName = Empty Then Exit Sub
            txtFileName.Text = .FileName
            txtOut.Text = EncryptFile(.FileName, txtPassword.Text, txtIn.Text, CBool(CheckTags.value), _
            cmbEncryption.ItemData(cmbEncryption.ListIndex), cmbLength.ItemData(cmbLength.ListIndex), _
                                    cmbBase.ItemData(cmbBase.ListIndex))
            lblCRC.Caption = GetCRC32(.FileName)
            txtIn.Text = Empty
        End With
    End If
    lblHash.Caption = objCap.GetHash(txtOut.Text, cmbAlgorithm.ItemData(cmbAlgorithm.ListIndex))
    Set objCap = Nothing
Exit Sub
ErrorCanc:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub


Private Sub cmdEncryptFile_Click()
    CipherFile Encrypt
End Sub

Private Sub cmdEncryptFileEasy_Click()
    CipherFileEasy Encrypt
End Sub

Private Sub cmdFind_Click()
    Dim cert As Certificate
    Dim FindType As CAPICOM_CERTIFICATE_FIND_TYPE
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim itmX As ListItem
    Dim StoreName As String
    Dim st As New Store
    Dim certs As New Certificates
    On Error GoTo ErrCert
    cmdFind.Enabled = False
     Select Case cmbStoreLocation.List(cmbStoreLocation.ListIndex)
         Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
           StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
         Case "CAPICOM_CURRENT_USER_STORE"
            StoreLocation = CAPICOM_CURRENT_USER_STORE
         Case "CAPICOM_LOCAL_MACHINE_STORE"
            StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
         Case "CAPICOM_SMART_CARD_USER_STORE"
            StoreLocation = CAPICOM_SMART_CARD_USER_STORE
         Case Else
                MsgBox "Please enter a valid Store location!", vbExclamation, App.Title
                cmdFind.Enabled = True
            Exit Sub
     End Select
    lstFoundCerts.ListItems.Clear
    StoreName = cmbStoreName.List(cmbStoreName.ListIndex)
    FindType = cmbFindType.ListIndex
    st.Open StoreLocation, StoreName
    Set certs = st.Certificates
    If FindType = CAPICOM_CERTIFICATE_FIND_TIME_EXPIRED Or _
       FindType = CAPICOM_CERTIFICATE_FIND_TIME_NOT_YET_VALID Or _
       FindType = CAPICOM_CERTIFICATE_FIND_TIME_VALID Then
        If txtCriteria.Text = Empty Then
            txtCriteria.Text = Now
        Else
            txtCriteria.Text = CDate(txtCriteria.Text)
        End If
        Set certs = certs.Find(FindType, txtCriteria.Text)
    Else
        If txtCriteria.Text <> "" Then
            Set certs = certs.Find(FindType, txtCriteria.Text)
        End If
    End If
    If certs.count = 0 Then
            MsgBox "No certificate found!" & vbCr & "Select oter criteria, please!", vbExclamation, App.Title
            cmdFind.Enabled = True
        Exit Sub
    End If
    For Each cert In certs
        Set itmX = lstFoundCerts.ListItems.Add(, , cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME))
        itmX.SubItems(1) = cert.GetInfo(CAPICOM_CERT_INFO_ISSUER_SIMPLE_NAME)
        itmX.SubItems(2) = cert.ValidFromDate
        itmX.SubItems(3) = cert.ValidToDate
    Next cert
    Set certs = Nothing
    cmdFind.Enabled = True
Exit Sub
ErrCert:
    MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
        cmdFind.Enabled = True
    Err.Clear
End Sub

Private Sub cmdFolderBrowse_Click()
    Dim strFolder As String
    strFolder = BrowseFolder("Save Certificate in:", App.Path)
    If strFolder <> "" And strFolder <> "Error!" Then
        If MakeDirectory(strFolder + "\" + "Certificates") Then
            txtDestPath.Text = strFolder + "\Certificates\"
        Else
            txtDestPath.Text = strFolder + "\"
        End If
    ' .... Save default path
    INI.DeleteKey "SETTING", "CERTIFICATE_PATH"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_PATH", txtDestPath.Text
    End If
End Sub

Private Sub cmdGenKey_Click()
    Set objGUIDE = New clsGUIDGenerator
    ts(5).Text = objGUIDE.CreateGUID("")
    Set objGUIDE = Nothing
End Sub

Private Sub cmdGetSign_Click()
    On Local Error GoTo ErrorGetSign
    SignedCode.FileName = txtFileToSing
    If SignedCode.FileName = Empty Then
            MsgBox "Select the file first!", vbExclamation, App.Title
        Exit Sub
    End If
    txtDescription.Text = SignedCode.Description
    txtDescriptionURL.Text = SignedCode.DescriptionURL
Exit Sub
ErrorGetSign:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbInformation, App.Title
    Err.Clear
End Sub



Private Sub cmdMakCab_Click()
    Dim intFile As Integer
    Dim i As Integer
    Dim sSetup As String
    Dim makeCabPath As String
    On Local Error GoTo ErrorCreate
    makeCabPath = App.Path + "\tools\Makecab.exe"
    sSetup = "_filelist.ddf"
    ' .... Now create the File sample.cab
    ' .... NOTE: Before call {Makecab.exe} make the file {_filelist.ddf} and fill the list of Files
    ' .... because {Makecab.exe} not support add multyple files ;)
    intFile = FreeFile()
    ' .... Step-1)
    Open txtDestPath + "_filelist.ddf" For Output As #intFile
        Print #intFile, ".Option Explicit"
        Print #intFile, ".Set CabinetNameTemplate=_sample.cab"
        Print #intFile, ".Set CompressionType=MSZIP"
        Print #intFile, ".Set Cabinet=on"
        Print #intFile, ";** List of Files"
        Print #intFile, ";*******************"
        For i = 0 To lstList.ListCount
            If lstList.List(i) <> Empty Then Print #intFile, """"; lstList.List(i); """"
        Next i
    Close #intFile
    ' .... Step-2)
    Open txtDestPath + "make_cab.bat" For Output As #intFile
        Print #intFile, """"; makeCabPath; """" & " /f " & """"; txtDestPath + sSetup; """"
        Print #intFile, "exit"
    Close #intFile
    
    ' .... Create CAB file
    If ShelledAPP(txtDestPath + "make_cab.bat", vbHide) = "End" Then
        MsgBox "File CAB process terminate success! Ok to Continue!", vbInformation, App.Title
    Else
        MsgBox "Batch process terminate with Error!", vbExclamation, App.Title
    ' .... Error? Delete batch file
            If Dir$(txtDestPath + "make_cab.bat") <> Empty Then Call Kill(txtDestPath + "make_cab.bat")
        Exit Sub
    End If
    
    ' .... open folder
    ShellExecute 0&, vbNullString, txtDestPath, vbNullString, "C:\", sShell.vbNormalFocus
Exit Sub
ErrorCreate:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub cmdNewEmail_Click()
    On Local Error Resume Next
    txtCC.Text = Empty
    txtTo.Text = Empty
    txtSubject.Text = Empty
    txtMessageBody.Text = Empty
    txtFromValue.Text = Empty
    txtToValue.Text = Empty
    txtCCValue.Text = Empty
    txtSubjectValue.Text = Empty
    txtTo.SetFocus
End Sub

Private Sub cmdOpenFolder_Click()
    On Local Error Resume Next
    ShellExecute 0&, vbNullString, txtDestPath.Text, vbNullString, "C:\", sShell.vbNormalFocus
End Sub


Private Sub cmdOpenMessage_Click()
    ' .... Show the open dialog
    On Local Error Resume Next
    With cDialog
        .CancelError = False
        .DefaultExt = ".eml"
        .Filter = "Message (*.eml;*.msg)|*.eml;*.msg|All Files (*.*)|*.*"
        .DialogTitle = "Select Message to Open:"
        .flags = (cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist)
        .ShowOpen
        If .FileName <> "" Then LoadMessage (.FileName)
    End With
End Sub

Private Sub cmdPicTools_Click()
    If cmdPicTools.Caption = "Base64 >>" Then
        cmdPicTools.Caption = "<< Base64"
        picTools.Visible = True
    Else
        cmdPicTools.Caption = "Base64 >>"
        picTools.Visible = False
    End If
End Sub

Private Sub cmdRandom_Click()
    Set objGUIDE = New clsGUIDGenerator
    txtPassword.Text = objGUIDE.CreateGUID("")
    Set objGUIDE = Nothing
End Sub

Private Sub cmdRemove_Click()
    On Local Error Resume Next
    If lstList.ListCount < 1 Then Exit Sub
    If lstList.List(lstList.ListIndex) <> txtDestPath.Text + "_setup.xml" Then _
    lstList.RemoveItem lstList.ListIndex
End Sub

Private Sub cmdResolveName_Click()
    Dim szNames As String
        If Len(Me.txtTo) > 0 Then
             szNames = Me.txtTo
             Call ResolveNames(szNames)
             Me.txtTo = szNames
         End If
         If Len(Me.txtCC) > 0 Then
             szNames = Me.txtCC
             Call ResolveNames(szNames)
             Me.txtCC = szNames
         End If
End Sub

Private Sub cmdSaveMessage_Click()
    Dim oBodyPart As CDO.IBodyPart
    Dim cFields As ADODB.Fields
    Dim oStream As ADODB.stream
    Dim oUtilities As New CAPICOM.Utilities
    Dim oRecipients As CAPICOM.Certificates
    Dim szNames As String
    
    On Error GoTo ErrorHandler
    
    ' .... Lets resolve the names to email addresses
    If Len(Me.txtTo) > 0 Then
        szNames = Me.txtTo
        Call ResolveNames(szNames)
        Me.txtTo = szNames
    End If
    If Len(Me.txtCC) > 0 Then
        szNames = Me.txtCC
        Call ResolveNames(szNames)
        Me.txtCC = szNames
    End If
    
    ' .... Gather the recipient certificates
    szNames = Me.txtTo & ";" & Me.txtCC
    Set oRecipients = ResolveNames(szNames)
    

    ' .... Make sure the minimum fields are populated
    If Me.txtTo.Text = "" Then
            MsgBox "You must specify at least one valid recipient in the 'To:' field!", vbExclamation, App.Title
        Exit Sub
    End If
    If Me.txtSubject.Text = "" Then
            MsgBox "You must specify a subject for the message in the 'Subject:' field!", vbExclamation, App.Title
        Exit Sub
    End If
    
    ' .... Create the message itself, this essentialy consists of setting a few header values and adding
    ' .... a new plain/text bodypart
    
    ' set sender, recipient, and subject.
    Set oMessage = New CDO.Message
    oMessage.To = Me.txtTo.Text
    oMessage.CC = Me.txtCC.Text
    oMessage.subject = Me.txtSubject.Text
    oMessage.Fields("urn:schemas:mailheader:date").value = oUtilities.LocalTimeToUTCTime(Now)
    oMessage.Fields.Update
    
    
    ' .... Set the current users email address to our best guess, we will get an authenticated address
    ' .... when we do signed mail from the signers certificate
    Dim szUserName As String
    If GetLoggedInUser(szUserName) Then
        oMessage.From = LCase(szUserName) & "@" & txtProvider.Text
    Else
        oMessage.From = "Anonymous"
    End If
    
    Set oBodyPart = oMessage.BodyPart.AddBodyPart
    Set cFields = oBodyPart.Fields
    cFields.Item(cdoContentType) = cdoTextPlain
    cFields.Update
    
    Set oStream = oBodyPart.GetDecodedContentStream
    oStream.WriteText txtMessageBody.Text
    oStream.Flush

    ' .... Sign, encrypt or sign/encrypt the message
    If ((CheckSign.value = 1) And (CheckEncrypt.value = 0)) Then
        ' .... It is a signed message
        If SignMessage(oMessage, True) = False Then
                MsgBox "Error to Sign E-Mail!", vbExclamation, App.Title
            Exit Sub
        End If
    ElseIf ((CheckSign.value = 0) And (CheckEncrypt.value = 1)) Then
        ' .... It is a encrypted message
        If EncryptMessage(oMessage, oRecipients) = False Then
                MsgBox "Error to Encrypt Message!", vbExclamation, App.Title
            Exit Sub
        End If
    ElseIf ((CheckSign.value = 1) And (CheckEncrypt.value = 1)) Then
        ' .... It is a signed and encrypted message
        If SignMessage(oMessage, True) = False Then
                MsgBox "Error to Sign E-Mail!", vbExclamation, App.Title
            Exit Sub
        End If
        If EncryptMessage(oMessage, oRecipients) = False Then
                MsgBox "Error to Encrypt Message!", vbExclamation, App.Title
            Exit Sub
        End If
    End If
    
    ' .... The message should look okay now, where would they like to save it to?
    With cDialog
        .CancelError = False
        .DefaultExt = ".eml"
        .Filter = "Message (*.eml;*.msg)|*.eml;*.msg|All Files (*.*)|*.*"
        .DialogTitle = "Select file to Save message to:"
        .flags = (cdlOFNCreatePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist)
        .ShowOpen
        If (.FileName <> "") Then
            oMessage.GetStream.SaveToFile .FileName, adSaveCreateOverWrite
        End If
    End With
Exit Sub
ErrorHandler:
   MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title

CleanUp:
    ' .... Clean up
    Set oBodyPart = Nothing
    Set cFields = Nothing
    Set oStream = Nothing
    Set oUtilities = Nothing
    Set oRecipients = Nothing
End Sub


Private Sub cmdSelect_Click()
    On Local Error GoTo ErrorSelect
    With cDialog
        .CancelError = True
        .DialogTitle = "Open file to Sign:"
        .Filter = "File  (*.exe)|*.exe|All Files (*.*)|*.*"
        .DefaultExt = ".exe"
        .ShowOpen
        If .FileName = Empty Then Exit Sub
        txtFileToSing.Text = .FileName
        SignedCode.FileName = txtFileToSing.Text
        INI.DeleteKey "CERTIFICATE", "FILE_TO_SIGN"
        INI.CreateKeyValue "CERTIFICATE", "FILE_TO_SIGN", .FileName
    End With
Exit Sub
ErrorSelect:
    If Err.Number = 32755 Then Exit Sub
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdShowCert_Click()
    On Local Error Resume Next
    If Not SignedCode Is Nothing Then
        If Not SignedCode.TimeStamper Is Nothing Then
            SignedCode.TimeStamper.Certificate.Display
        Else
            MsgBox "The file hasn't been Time Stamped!", vbInformation, App.Title
        End If
    Else
           MsgBox "Please select a File first to Sign!", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdSignAndVerify_Click()
    SignAndVerify
End Sub

Private Sub cmdSignDistortVerify_Click()
    SignDistortAndVerify
End Sub


Private Sub cmdSingFile_Click()
    Dim cert As Certificate
    Dim certs As New Certificates
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim StoreName As String
    Dim st As New Store
    Dim i As Integer
    On Error GoTo ErrorSign
    
        If SignedCode.FileName = Empty Then
                MsgBox "Select the file first!", vbExclamation, App.Title
            Exit Sub
        End If
        
        If txtCer(1).Text = Empty Then
                MsgBox "Select Certificate first!", vbExclamation, App.Title
            Exit Sub
        End If
        
        Select Case txtCer(2).Text
            Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
                StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
            Case "CAPICOM_CURRENT_USER_STORE"
                StoreLocation = CAPICOM_CURRENT_USER_STORE
            Case "CAPICOM_LOCAL_MACHINE_STORE"
                StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
            Case "CAPICOM_SMART_CARD_USER_STORE"
                StoreLocation = CAPICOM_SMART_CARD_USER_STORE
            Case Else
                MsgBox "Select Certificate first!", vbExclamation, App.Title
        End Select
        
        StoreName = txtCer(1).Text
        st.Open StoreLocation, StoreName, CAPICOM_STORE_OPEN_READ_ONLY
        Set certs = st.Certificates
        Signer.Certificate = st.Certificates.Item(TrovaCertificato(st))
        
        SignedCode.Description = txtDescription.Text
        SignedCode.DescriptionURL = txtDescriptionURL.Text
        SignedCode.FileName = txtFileToSing.Text
        SignedCode.Sign Signer
        MsgBox "The fle has been signed Succesfully!", vbInformation, App.Title
Exit Sub
ErrorSign:
    If Err.Number <> CAPICOM_E_CANCELLED Then
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbInformation, App.Title
    End If
    Err.Clear
End Sub

Private Sub cmdTimeStamp_Click()
    On Error GoTo ErrorTimeStamp
        SignedCode.FileName = txtFileToSing
        If SignedCode.FileName = Empty Then
                MsgBox "Select the file first!", vbExclamation, App.Title
            Exit Sub
        End If
        SignedCode.TimeStamp URL
        MsgBox "The file has been Time Stamped with:" & vbCr & URL, vbInformation, App.Title
Exit Sub
ErrorTimeStamp:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbInformation, App.Title
    Err.Clear
End Sub

Private Sub cmdVerify_Click()
    On Error GoTo ErrorVerify
    SignedCode.FileName = txtFileToSing
        If SignedCode.FileName = Empty Then
                MsgBox "Select the file to Verify!", vbExclamation, App.Title
            Exit Sub
        End If
        SignedCode.Verify
        MsgBox "The file has been verified!", vbInformation, App.Title
Exit Sub
ErrorVerify:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbInformation, App.Title
    Err.Clear
End Sub

Private Sub Form_Initialize()
    ' .... Display the Copyright
    Me.Caption = tTitle & " 2008/" & Format(Now, "yyyy") & "  © Salvo Cortesiano!"
    lblInfo.Caption = "© 2008/" & Format(Now, "mm") & "," & Format(Now, "yyyy") & " by Salvo Cortesiano. All Right Reserved!"
    ' .... Reset INI File Path
    INI.ResetINIFilePath
    ' .... CF Path
    cfPath = App.Path + "\CF\CFC.exe"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    On Local Error GoTo ErrorLoad
    ' .... Read value
    ssTop = INI.GetKeyValue("FORM", "S_TOP")
    ssLeft = INI.GetKeyValue("FORM", "S_LEFT")
    ' .... Position MainForm
    If Len(ssLeft) = 0 Then
        ' .... Center Form
        ssTop = (Screen.Height - frmMain.Height) \ 2
        ssLeft = (Screen.Width - frmMain.Width) \ 2
        frmMain.Move ssLeft, ssTop
    Else
        frmMain.Move ssLeft, ssTop
    End If
    ' .... HASH
    If INI.GetKeyValue("SETTING", "ENCODE_HASH") <> "" Then
        If cmbAlgorithm.ListIndex <> Empty Then cmbAlgorithm.ListIndex = INI.GetKeyValue("SETTING", "ENCODE_HASH")
    Else
        If cmbAlgorithm.ListIndex <> Empty Then cmbAlgorithm.ListIndex = 6
    End If
    ' .... ALGORITHM
    If INI.GetKeyValue("SETTING", "ENCODE_ALGORITHM") <> "" Then
        If cmbEncryption.ListIndex <> Empty Then cmbEncryption.ListIndex = INI.GetKeyValue("SETTING", "ENCODE_ALGORITHM")
    Else
        If cmbEncryption.ListIndex <> Empty Then cmbEncryption.ListIndex = 1
    End If
    ' ....LENGTH
    If INI.GetKeyValue("SETTING", "ENCODE_LENGTH") <> "" Then
        If cmbLength.ListIndex <> Empty Then cmbLength.ListIndex = INI.GetKeyValue("SETTING", "ENCODE_LENGTH")
    Else
        If cmbLength.ListIndex <> Empty Then cmbLength.ListIndex = 5
    End If
    ' .... BASE
    If INI.GetKeyValue("SETTING", "ENCODE_BASE") <> "" Then
        If cmbBase.ListIndex <> Empty Then cmbBase.ListIndex = INI.GetKeyValue("SETTING", "ENCODE_BASE")
    Else
        If cmbBase.ListIndex <> Empty Then cmbBase.ListIndex = 0
    End If
    ' ..... ENCODE to FILE
    If INI.GetKeyValue("SETTING", "ENCODE_TO_FILE") <> "" Then _
    CheckToFile.value = INI.GetKeyValue("SETTING", "ENCODE_TO_FILE") Else _
    CheckToFile.value = 0
    ' .... Write TAGs
    If INI.GetKeyValue("SETTING", "INCLUDE_TAGs") <> "" Then _
    CheckTags.value = INI.GetKeyValue("SETTING", "INCLUDE_TAGs") Else _
    CheckTags.value = 0
    ' .... Show PassWord?
    If INI.GetKeyValue("SETTING", "SHOW_PASSWORD") <> "" Then _
    CheckShowPsW.value = INI.GetKeyValue("SETTING", "SHOW_PASSWORD") Else _
    CheckShowPsW.value = 1
    ' .... Last Encoded File
    If INI.GetKeyValue("SETTING", "LAST_ENCODE_FILEPATH") <> "" Then _
    txtFileName.Text = INI.GetKeyValue("SETTING", "LAST_ENCODE_FILEPATH")
    ' .... PassWord
    If INI.GetKeyValue("SETTING", "ENCODE_PASSWORD") <> "" Then _
    txtPassword.Text = INI.GetKeyValue("SETTING", "ENCODE_PASSWORD")
    For i = 0 To 7
        If INI.GetKeyValue("SETTING", "SIGNATORY_" & i) <> "" Then _
        ts(i).Text = INI.GetKeyValue("SETTING", "SIGNATORY_" & i)
    Next i
    ' .... Create PFX File
    If INI.GetKeyValue("SETTING", "CREATE_ALSO_PFX") = Empty Then _
    CheckPfx.value = 0 Else CheckPfx.value = INI.GetKeyValue("SETTING", "CREATE_ALSO_PFX")
    
    i = 0
    
    For i = 0 To 6
        If INI.GetKeyValue("SETTING", "CERTIFICATE_USE" & i) <> "" Then _
        CheckUsers(i).value = INI.GetKeyValue("SETTING", "CERTIFICATE_USE" & i) _
        Else CheckUsers(i).value = 0
    Next i
    
    txtKeyPassWord.Text = INI.GetKeyValue("SETTING", "SIGNATORY_PASSWORD")
    
    If INI.GetKeyValue("CERTIFICATE", "STORE") <> "" Then
        If cmbStoreName.ListIndex <> Empty Then cmbStoreName.ListIndex = INI.GetKeyValue("CERTIFICATE", "STORE")
    Else
        If cmbStoreName.ListIndex <> Empty Then cmbStoreName.ListIndex = 0
    End If
    If INI.GetKeyValue("CERTIFICATE", "STORE_LOCATION") <> "" Then
        If cmbStoreLocation.ListIndex <> Empty Then cmbStoreLocation.ListIndex = INI.GetKeyValue("CERTIFICATE", "STORE_LOCATION")
    Else
        If cmbStoreLocation.ListIndex <> Empty Then cmbStoreLocation.ListIndex = 1
    End If
    If INI.GetKeyValue("CERTIFICATE", "STORE_TYPE") <> "" Then
        If cmbFindType.ListIndex <> Empty Then cmbFindType.ListIndex = INI.GetKeyValue("CERTIFICATE", "STORE_TYPE")
    Else
        If cmbFindType.ListIndex <> Empty Then cmbFindType.ListIndex = 1
    End If
    cmdFind = True
    
    Call GetCert
    
    txtFileToSing.Text = INI.GetKeyValue("CERTIFICATE", "FILE_TO_SIGN")
    txtDestPath.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_PATH")
    
    txtAutority.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_AUTORITY_NAME")
    txtCertificateName.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_NAME")
    
    txtValidFrom.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_VALID_FROM")
    txtValidTo.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_VALID_TO")
    txtPassCert.Text = INI.GetKeyValue("SETTING", "CERTIFICATE_PASSWORD")
    
    If txtValidFrom.Text <> Empty Then _
        txtValidFrom.ToolTipText = Format(txtValidFrom.Text, "Long Date")
    If txtValidTo.Text <> Empty Then _
        txtValidTo.ToolTipText = Format(txtValidTo.Text, "Long Date")
    
    CheckPutMyValidity.value = INI.GetKeyValue("SETTING", "INCLUDE_VALIDITY")
    
    Call GetAllIP
    
    If txtInternalIPs.ListCount > 0 Then
        If INI.GetKeyValue("SETTING", "IPS") <> Empty Or INI.GetKeyValue("SETTING", "IPS") <> "-1" Then
            txtInternalIPs.ListIndex = INI.GetKeyValue("SETTING", "IPS")
        End If
    End If
    
    If CheckIP.Enabled = True Then _
    CheckIP.value = INI.GetKeyValue("SETTING", "INCLUDE_IP")
    
    CheckInstall.value = INI.GetKeyValue("SETTING", "CREATE_AND_INSTALL")
    
    If INI.GetKeyValue("SETTING", "CERTIFICATE_STORES") <> Empty Or INI.GetKeyValue("SETTING", "CERTIFICATE_STORES") <> "-1" Then _
    cmb_Stores.ListIndex = INI.GetKeyValue("SETTING", "CERTIFICATE_STORES") Else cmb_Stores.ListIndex = 0
    
    txtInternalIPs.AddItem Compuer_Name
    txtInternalIPs.AddItem GetUser_Name
    
    cmbsStoresList.ListIndex = 3
Exit Sub
ErrorLoad:
    Err.Clear
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure to Close this Application?", vbYesNo + vbInformation + _
        vbDefaultButton2, App.Title) = vbYes Then
        ' .... Close = True
        readyToClose = True
        ' .... Save Setting to File *.INI
        SaveSettings
        ' .... Release Class INI
        Set INI = Nothing
        ' .... Release the Class GUIDE
        Set objGUIDE = Nothing
        ' .... Release the Class CAPICOM
        Set objCap = Nothing
        Else
            readyToClose = False
    End If
        Cancel = Not readyToClose
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' .... Sure to close Program?
    End
End Sub



Private Function ReadEncrypedFile(ByVal sFilePath As String, ByVal sPassword As String, _
                                    Optional sDelimiter As String = "|") As String
    Dim sContents As String
    Dim sAr() As String
    Dim i As Long
    Dim sTmp As String
    Dim sItem As String
    On Local Error GoTo ErrorDecrypt
    ' .... Read file and formatted contents
    sContents = DecryptFileFromFile(sFilePath, sPassword)
    
    i = 0
    sTmp = ""
    
    sAr = Split(sContents, sDelimiter)
    For i = 0 To UBound(sAr)
        sItem = sAr(i)
        If InStr(sItem, "Signatory") Then
            sTmp = sTmp & "Signatory Name = " & Replace(sItem, "Signatory", "") & vbCrLf
        ElseIf InStr(sItem, "Certifier") Then
            sTmp = sTmp & "Certifier = " & Replace(sItem, "Certifier", "") & vbCrLf
        ElseIf InStr(sItem, "CF") Then
            sTmp = sTmp & "Code F = " & Replace(sItem, "CF", "") & vbCrLf
        ElseIf InStr(sItem, "State") Then
            sTmp = sTmp & "State = " & Replace(sItem, "State", "") & vbCrLf
        ElseIf InStr(sItem, "SignatureID") Then
            sTmp = sTmp & "ID = " & Replace(sItem, "SignatureID", "") & vbCrLf
        ElseIf InStr(sItem, "ValidityDate") Then
            sTmp = sTmp & "Validity = " & Replace(sItem, "ValidityDate", "") & vbCrLf
        ElseIf InStr(sItem, "Name") Then
            sTmp = sTmp & "Release To = " & Replace(sItem, "Name", "") & vbCrLf
        ElseIf InStr(sItem, "Note") Then
            sTmp = sTmp & "Note = " & Replace(sItem, "Note", "") & vbCrLf
        End If
    Next i
    
    ReadEncrypedFile = sTmp
Exit Function
ErrorDecrypt:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Private Sub SaveSettings()
    Dim i As Integer
    On Local Error Resume Next
    
    ' .... Signatory Fields
    For i = 0 To 7
        INI.DeleteKey "SETTING", "SIGNATORY_" & i
        INI.CreateKeyValue "SETTING", "SIGNATORY_" & i, ts(i).Text
    Next i
    
    INI.DeleteKey "SETTING", "SIGNATORY_PASSWORD"
    INI.CreateKeyValue "SETTING", "SIGNATORY_PASSWORD", txtKeyPassWord.Text
    
    INI.DeleteKey "SETTING", "ENCODE_TO_FILE"
    INI.CreateKeyValue "SETTING", "ENCODE_TO_FILE", CheckToFile.value
    
    INI.DeleteKey "SETTING", "INCLUDE_TAGs"
    INI.CreateKeyValue "SETTING", "INCLUDE_TAGs", CheckTags.value
    
    INI.DeleteKey "SETTING", "SHOW_PASSWORD"
    INI.CreateKeyValue "SETTING", "SHOW_PASSWORD", CheckShowPsW.value
    
    INI.DeleteKey "SETTING", "ENCODE_HASH"
    INI.CreateKeyValue "SETTING", "ENCODE_HASH", cmbAlgorithm.ListIndex
    
    INI.DeleteKey "SETTING", "ENCODE_ALGORITHM"
    INI.CreateKeyValue "SETTING", "ENCODE_ALGORITHM", cmbEncryption.ListIndex
    
    INI.DeleteKey "SETTING", "ENCODE_LENGTH"
    INI.CreateKeyValue "SETTING", "ENCODE_LENGTH", cmbLength.ListIndex
    
    INI.DeleteKey "SETTING", "ENCODE_BASE"
    INI.CreateKeyValue "SETTING", "ENCODE_BASE", cmbBase.ListIndex
    
    INI.DeleteKey "SETTING", "LAST_ENCODE_FILEPATH"
    If txtFileName.Text <> Empty Then INI.CreateKeyValue "SETTING", "LAST_ENCODE_FILEPATH", txtFileName.Text
    
    INI.DeleteKey "SETTING", "ENCODE_PASSWORD"
    If txtPassword.Text <> Empty Then INI.CreateKeyValue "SETTING", "ENCODE_PASSWORD", txtPassword.Text
    
    INI.DeleteKey "CERTIFICATE", "STORE"
    INI.CreateKeyValue "CERTIFICATE", "STORE", cmbStoreName.ListIndex
    INI.DeleteKey "CERTIFICATE", "STORE_LOCATION"
    INI.CreateKeyValue "CERTIFICATE", "STORE_LOCATION", cmbStoreLocation.ListIndex
    INI.DeleteKey "CERTIFICATE", "STORE_TYPE"
    INI.CreateKeyValue "CERTIFICATE", "STORE_TYPE", cmbFindType.ListIndex
    
    INI.DeleteKey "SETTING", "CREATE_ALSO_PFX"
    INI.CreateKeyValue "SETTING", "CREATE_ALSO_PFX", CheckPfx.value
    
    i = 0
    
    For i = 0 To 6
        INI.DeleteKey "SETTING", "CERTIFICATE_USE" & i
        INI.CreateKeyValue "SETTING", "CERTIFICATE_USE" & i, CheckUsers(i).value
    Next i
    
    If txtInternalIPs.ListCount > 0 Then
        INI.DeleteKey "SETTING", "IPS"
        INI.CreateKeyValue "SETTING", "IPS", txtInternalIPs.ListIndex
    End If
    
    INI.DeleteKey "SETTING", "INCLUDE_IP"
    INI.CreateKeyValue "SETTING", "INCLUDE_IP", CheckIP.value
    
    INI.DeleteKey "SETTING", "INCLUDE_VALIDITY"
    INI.CreateKeyValue "SETTING", "INCLUDE_VALIDITY", CheckPutMyValidity.value
    
    INI.DeleteKey "SETTING", "CREATE_AND_INSTALL"
    INI.CreateKeyValue "SETTING", "CREATE_AND_INSTALL", CheckInstall.value
    
    INI.DeleteKey "SETTING", "CERTIFICATE_STORES"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_STORES", cmb_Stores.ListIndex
    
    ' .... Position form
    If Me.WindowState <> vbMinimized Then
        INI.DeleteKey "FORM", "S_LEFT"
        INI.CreateKeyValue "FORM", "S_LEFT", frmMain.Left
        INI.DeleteKey "FORM", "S_TOP"
        INI.CreateKeyValue "FORM", "S_TOP", frmMain.Top
    End If
End Sub

Private Sub imgCF_Click()
    If Dir$(cfPath) = Empty Then
            MsgBox "CFC.exe not Found in {" & App.Path + "\CF" & "}.", vbExclamation, App.Title
        Exit Sub
    End If
    Clipboard.Clear
    If ShelledAPP(cfPath) = "End" Then
        If Clipboard.GetFormat(vbCFText) Then
            ts(2).Text = Clipboard.GetText()
        Else
            MsgBox "No data to Paste into Clipboard!", vbExclamation, App.Title
        End If
    Else
        MsgBox "Error into Function() Shelled!", vbExclamation, App.Title
    End If
End Sub


Private Sub lstFoundCerts_Click()
    If Not lstFoundCerts.SelectedItem Is Nothing Then
        If InfoCertificate(lstFoundCerts.SelectedItem.Text) Then:
    End If
End Sub

Private Sub lstFoundCerts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lstFoundCerts.SelectedItem Is Nothing And Button = 2 Then
        mnuCert(0).Caption = "Show {" & lstFoundCerts.SelectedItem.Text & "}"
        mnuCert(1).Caption = "Delete {" & lstFoundCerts.SelectedItem.Text & "}"
        mnuCert(3).Caption = "Use this {" & lstFoundCerts.SelectedItem.Text & "} for Sign!"
        PopupMenu mnuCertificate
    End If
End Sub


Private Sub mnuCert_Click(Index As Integer)
    Select Case Index
        Case 0 ' Show Certificate
            ShowCertificate lstFoundCerts.SelectedItem.Text
        Case 1 ' Delete Certificate
            If MsgBox("Are you sure to Remuve this Certificate?" & vbCr & vbCr & lstFoundCerts.SelectedItem.Text, vbYesNo + vbInformation + _
                vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                
        Case 3 ' Use Certificate for Sign
            If MsgBox("Are you sure to Use this Certificate for Sign?" & vbCr & vbCr & lstFoundCerts.SelectedItem.Text, vbYesNo + vbInformation + _
                vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            ' Save To INI
            INI.DeleteKey "CERTIFICATE", "NAME"
            INI.CreateKeyValue "CERTIFICATE", "NAME", lstFoundCerts.SelectedItem.Text
            INI.DeleteKey "CERTIFICATE", "STORE"
            INI.CreateKeyValue "CERTIFICATE", "STORE_", cmbStoreName.List(cmbStoreName.ListIndex)
            INI.DeleteKey "CERTIFICATE", "STORE_LOCATION_"
            INI.CreateKeyValue "CERTIFICATE", "STORE_LOCATION_", cmbStoreLocation.List(cmbStoreLocation.ListIndex)
            INI.DeleteKey "CERTIFICATE", "STORE_TYPE_"
            INI.CreateKeyValue "CERTIFICATE", "STORE_TYPE_", cmbFindType.List(cmbFindType.ListIndex)
            Call GetCert
    End Select
End Sub

Private Sub TabCrypt_Click()
    Select Case TabCrypt.SelectedItem.Key
        Case "Providers"
            picCrypt(0).Visible = True
            picCrypt(2).Visible = False
            picCrypt(1).Visible = False
        Case "Hash"
            picCrypt(0).Visible = False
            picCrypt(2).Visible = False
            picCrypt(1).Visible = True
        Case "Keys"
            picCrypt(0).Visible = False
            picCrypt(1).Visible = False
            picCrypt(2).Visible = True
    End Select
End Sub

Private Sub TBS_Click()
On Local Error GoTo ErrorHandler
    If tbsKey = TBS.SelectedItem.Key Then Exit Sub
    Select Case TBS.SelectedItem.Key
        Case "EncDec"
            picsTabs(6).Visible = False
            picsTabs(5).Visible = False
            picsTabs(4).Visible = False
            picsTabs(3).Visible = False
            picsTabs(2).Visible = False
            picsTabs(0).Visible = True
            picsTabs(1).Visible = False
        Case "SerialKey"
            picsTabs(6).Visible = False
            picsTabs(5).Visible = False
            picsTabs(4).Visible = False
            picsTabs(3).Visible = False
            picsTabs(2).Visible = False
            picsTabs(1).Visible = True
            picsTabs(0).Visible = False
        Case "Certificate"
            picsTabs(6).Visible = False
            picsTabs(5).Visible = False
            picsTabs(4).Visible = False
            picsTabs(3).Visible = False
            picsTabs(2).Visible = True
            picsTabs(1).Visible = False
            picsTabs(0).Visible = False
            lstFoundCerts.SetFocus
            lstFoundCerts.ListItems(1).Selected = True
            lstFoundCerts_Click
        Case "SignSoftware"
            picsTabs(6).Visible = False
            picsTabs(5).Visible = False
            picsTabs(4).Visible = False
            picsTabs(3).Visible = True
            picsTabs(2).Visible = False
            picsTabs(1).Visible = False
            picsTabs(0).Visible = False
            Call GetCert
            txtFileToSing.Text = INI.GetKeyValue("CERTIFICATE", "FILE_TO_SIGN")
            SignedCode.FileName = txtFileToSing.Text
        Case "CryptoAPI"
            LoadProviders
            CheckOS
            TabCrypt.TabIndex = 0
            picsTabs(6).Visible = False
            picsTabs(5).Visible = False
            picsTabs(4).Visible = True
            picsTabs(3).Visible = False
            picsTabs(2).Visible = False
            picsTabs(1).Visible = False
            picsTabs(0).Visible = False
        Case "CreateCertificate"
            picsTabs(6).Visible = False
            picsTabs(5).Visible = True
            picsTabs(4).Visible = False
            picsTabs(3).Visible = False
            picsTabs(2).Visible = False
            picsTabs(1).Visible = False
            picsTabs(0).Visible = False
        Case "SignMIME"
            picsTabs(6).Visible = True
            picsTabs(5).Visible = False
            picsTabs(4).Visible = False
            picsTabs(3).Visible = False
            picsTabs(2).Visible = False
            picsTabs(1).Visible = False
            picsTabs(0).Visible = False
    End Select
    tbsKey = TBS.SelectedItem.Key
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub
Private Sub ShowCertificate(sCert As String)
    Dim cert As Certificate
    Dim certs As New Certificates
    Dim FindType As CAPICOM_CERTIFICATE_FIND_TYPE
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim StoreName As String
    Dim st As New Store
    On Local Error GoTo ErrorTag
    If Not lstFoundCerts.SelectedItem Is Nothing Then
        Select Case cmbStoreLocation.List(cmbStoreLocation.ListIndex)
         Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
           StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
         Case "CAPICOM_CURRENT_USER_STORE"
            StoreLocation = CAPICOM_CURRENT_USER_STORE
         Case "CAPICOM_LOCAL_MACHINE_STORE"
            StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
         Case "CAPICOM_SMART_CARD_USER_STORE"
            StoreLocation = CAPICOM_SMART_CARD_USER_STORE
         Case Else
            Exit Sub
        End Select
        StoreName = cmbStoreName.List(cmbStoreName.ListIndex)
        FindType = cmbFindType.ListIndex
        st.Open StoreLocation, StoreName, CAPICOM_STORE_OPEN_READ_ONLY
        Set certs = st.Certificates
        If certs.count = 0 Then Exit Sub
        For Each cert In certs
            If LCase(sCert) = LCase(cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME)) Then
                    cert.Display
                Exit For
            End If
        Next cert
    End If
    Set certs = Nothing
Exit Sub
ErrorTag:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Function InfoCertificate(sCert As String) As Boolean
    Dim cert As Certificate
    Dim certs As New Certificates
    Dim FindType As CAPICOM_CERTIFICATE_FIND_TYPE
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim StoreName As String
    Dim st As New Store
    Dim sTxtTemp As String
    On Local Error GoTo ErrorTag
        Select Case cmbStoreLocation.List(cmbStoreLocation.ListIndex)
         Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
           StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
         Case "CAPICOM_CURRENT_USER_STORE"
            StoreLocation = CAPICOM_CURRENT_USER_STORE
         Case "CAPICOM_LOCAL_MACHINE_STORE"
            StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
         Case "CAPICOM_SMART_CARD_USER_STORE"
            StoreLocation = CAPICOM_SMART_CARD_USER_STORE
         Case Else
                InfoCertificate = False
            Exit Function
        End Select
        StoreName = cmbStoreName.List(cmbStoreName.ListIndex)
        FindType = cmbFindType.ListIndex
        st.Open StoreLocation, StoreName
        Set certs = st.Certificates
        If certs.count = 0 Then Exit Function
        txtInfCert.Text = Empty
        For Each cert In certs
            If LCase(sCert) = LCase(cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME)) Then
                    sTxtTemp = sTxtTemp & "DNS Name: " & cert.GetInfo(CAPICOM_CERT_INFO_ISSUER_DNS_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Issuer Email: " & cert.GetInfo(CAPICOM_CERT_INFO_ISSUER_EMAIL_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Issuer Name: " & cert.GetInfo(CAPICOM_CERT_INFO_ISSUER_SIMPLE_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Issuer UPN: " & cert.GetInfo(CAPICOM_CERT_INFO_ISSUER_UPN) & vbCrLf
                    sTxtTemp = sTxtTemp & "Subject DSN Name: " & cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_DNS_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Subject Email Name: " & cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Subject Name: " & cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME) & vbCrLf
                    sTxtTemp = sTxtTemp & "Subject UPN: " & cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_UPN) & vbCrLf
                    sTxtTemp = sTxtTemp & "Private Key: " & cert.HasPrivateKey & vbCrLf
                    sTxtTemp = sTxtTemp & "Valid EKU: " & cert.IsValid.EKU & vbCrLf
                    sTxtTemp = sTxtTemp & "Valid Flag: " & cert.IsValid.CheckFlag & vbCrLf
                    sTxtTemp = sTxtTemp & "Valid: " & cert.IsValid.Result & vbCrLf
                    ' Additional Info
                    sTxtTemp = sTxtTemp & "CRL Sign Enabled: " & cert.KeyUsage.IsCRLSignEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Critical: " & cert.KeyUsage.IsCritical & vbCrLf
                    sTxtTemp = sTxtTemp & "Data Encipherment Enabled: " & cert.KeyUsage.IsDataEnciphermentEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Decipher Only Enabled: " & cert.KeyUsage.IsDecipherOnlyEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Digital Signature Enabled: " & cert.KeyUsage.IsDigitalSignatureEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Encipher Only Enabled: " & cert.KeyUsage.IsEncipherOnlyEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Key Agreement Enabled: " & cert.KeyUsage.IsKeyAgreementEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Key CertSign Enabled: " & cert.KeyUsage.IsKeyCertSignEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Key Encipherment Enabled: " & cert.KeyUsage.IsKeyEnciphermentEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Non Repudiation Enabled: " & cert.KeyUsage.IsNonRepudiationEnabled & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Present: " & cert.KeyUsage.IsPresent & vbCrLf
                    ' Oter Info
                    On Error Resume Next ' Prevent the Crash if the PrivateKey do not exist ;)
                    sTxtTemp = sTxtTemp & "Private Key Name: " & cert.PrivateKey.ContainerName & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Accessible: " & cert.PrivateKey.IsAccessible & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Exportable: " & cert.PrivateKey.IsExportable & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Hardware Device: " & cert.PrivateKey.IsHardwareDevice & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Machine Keyset: " & cert.PrivateKey.IsMachineKeyset & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Protected: " & cert.PrivateKey.IsProtected & vbCrLf
                    sTxtTemp = sTxtTemp & "Is Removable: " & cert.PrivateKey.IsRemovable & vbCrLf
                    sTxtTemp = sTxtTemp & "Key Special: " & cert.PrivateKey.KeySpec & vbCrLf
                    sTxtTemp = sTxtTemp & "Provider Name: " & cert.PrivateKey.ProviderName & vbCrLf
                    sTxtTemp = sTxtTemp & "Provider Type: " & cert.PrivateKey.ProviderType & vbCrLf
                    sTxtTemp = sTxtTemp & "Unique Container Name: " & cert.PrivateKey.UniqueContainerName & vbCrLf
                    ' Oter Info 2
                    sTxtTemp = sTxtTemp & "Serial Number: " & cert.SerialNumber & vbCrLf
                    sTxtTemp = sTxtTemp & "P. Key Algorithm: " & cert.PublicKey.Algorithm & vbCrLf
                    sTxtTemp = sTxtTemp & "P. Key Algorithm FriendlyName: " & cert.PublicKey.Algorithm.FriendlyName & vbCrLf
                    sTxtTemp = sTxtTemp & "P. Key Algorithm Name: " & cert.PublicKey.Algorithm.Name & vbCrLf
                    sTxtTemp = sTxtTemp & "P. Key Algorithm Value: " & cert.PublicKey.Algorithm.value & vbCrLf
                    ' Oter Info 3
                    sTxtTemp = sTxtTemp & "Template is Critical: " & cert.Template.IsCritical & vbCrLf
                    sTxtTemp = sTxtTemp & "Template is Present: " & cert.Template.IsPresent & vbCrLf
                    sTxtTemp = sTxtTemp & "Major Version: " & cert.Template.MajorVersion & vbCrLf
                    sTxtTemp = sTxtTemp & "Minor Version: " & cert.Template.MinorVersion & vbCrLf
                    sTxtTemp = sTxtTemp & "Name: " & cert.Template.Name & vbCrLf
                    sTxtTemp = sTxtTemp & "OID: " & cert.Template.OID & vbCrLf
                    ' Oter Info 4
                    sTxtTemp = sTxtTemp & "Thumb Print: " & cert.Thumbprint & vbCrLf
                    sTxtTemp = sTxtTemp & "Valid From: " & cert.ValidFromDate & vbCrLf
                    sTxtTemp = sTxtTemp & "Valid To: " & cert.ValidToDate & vbCrLf
                    sTxtTemp = sTxtTemp & "Version: " & cert.Version
                Exit For
            End If
        Next cert
        txtInfCert.Text = sTxtTemp
        InfoCertificate = True
    Set certs = Nothing
Exit Function
ErrorTag:
    MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
        InfoCertificate = False
    Err.Clear
End Function

Private Function DeleteCertificate(sCert As String) As Boolean
    Dim cert As Certificate
    Dim certs As New Certificates
    Dim FindType As CAPICOM_CERTIFICATE_FIND_TYPE
    Dim StoreLocation As CAPICOM_STORE_LOCATION
    Dim StoreName As String
    Dim st As New Store
    Dim i As Long
    On Local Error GoTo ErrorTag
        Select Case cmbStoreLocation.List(cmbStoreLocation.ListIndex)
         Case "CAPICOM_ACTIVE_DIRECTORY_USER_STORE"
           StoreLocation = CAPICOM_ACTIVE_DIRECTORY_USER_STORE
         Case "CAPICOM_CURRENT_USER_STORE"
            StoreLocation = CAPICOM_CURRENT_USER_STORE
         Case "CAPICOM_LOCAL_MACHINE_STORE"
            StoreLocation = CAPICOM_LOCAL_MACHINE_STORE
         Case "CAPICOM_SMART_CARD_USER_STORE"
            StoreLocation = CAPICOM_SMART_CARD_USER_STORE
         Case Else
                DeleteCertificate = False
            Exit Function
        End Select
        StoreName = cmbStoreName.List(cmbStoreName.ListIndex)
        FindType = cmbFindType.ListIndex
        st.Open StoreLocation, StoreName
        Set certs = st.Certificates
        If certs.count = 0 Then Exit Function
        i = 1
        For Each cert In certs
            If LCase(sCert) = LCase(cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME)) Then
                    certs.Remove (i - 1)
                Exit For
            End If
            i = i + 1
        Next cert
        DeleteCertificate = True
        cmdFind = True
    Set certs = Nothing
Exit Function
ErrorTag:
    MsgBox "Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
        DeleteCertificate = False
    Err.Clear
End Function

Private Sub GetCert()
    On Local Error Resume Next
    txtCer(0).Text = INI.GetKeyValue("CERTIFICATE", "NAME")
    txtCer(1).Text = INI.GetKeyValue("CERTIFICATE", "STORE_")
    txtCer(2).Text = INI.GetKeyValue("CERTIFICATE", "STORE_LOCATION_")
    txtCer(3).Text = INI.GetKeyValue("CERTIFICATE", "STORE_TYPE_")
End Sub

Private Function TrovaCertificato(myStore As Store) As Integer
    Dim i As Integer
    Dim bolTrovato As Boolean
    On Local Error GoTo ErrorFound
    bolTrovato = False
    i = 1
    While (i <= myStore.Certificates.count) And Not (bolTrovato)
        bolTrovato = (myStore.Certificates.Item(i).GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME) = txtCer(0).Text)
        i = i + 1
    Wend
    If bolTrovato Then
        TrovaCertificato = i - 1
    Else
        TrovaCertificato = 0
    End If
Exit Function
ErrorFound:
        TrovaCertificato = 0
Err.Clear
End Function

Private Function GetStore() As Integer
    On Local Error GoTo ErrorGetStore
    If UCase$(txtCer(2).Text) = UCase$(CAPICOM_ACTIVE_DIRECTORY_USER_STORE) Then
        GetStore = 0
    ElseIf UCase$(txtCer(2).Text) = UCase$(CAPICOM_CURRENT_USER_STORE) Then
        GetStore = 1
    ElseIf UCase$(txtCer(2).Text) = UCase$(CAPICOM_LOCAL_MACHINE_STORE) Then
        GetStore = 2
    ElseIf UCase$(txtCer(2).Text) = UCase$(CAPICOM_MEMORY_STORE) Then
        GetStore = 3
    ElseIf UCase$(txtCer(2).Text) = UCase$(CAPICOM_SMART_CARD_USER_STORE) Then
        GetStore = 4
    Else
        GetStore = 1
    End If
Exit Function
ErrorGetStore:
        GetStore = 0
    Err.Clear
End Function

Private Sub txtAutority_Change()
    INI.DeleteKey "SETTING", "CERTIFICATE_AUTORITY_NAME"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_AUTORITY_NAME", txtAutority.Text
End Sub

Private Sub txtCertificateName_Change()
    INI.DeleteKey "SETTING", "CERTIFICATE_NAME_"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_NAME", txtCertificateName.Text
End Sub

Private Sub txtDescription_Change()
    SignedCode.Description = txtDescription.Text
End Sub

Private Sub txtDescriptionURL_Change()
    SignedCode.DescriptionURL = txtDescriptionURL.Text
End Sub



Private Sub GetProviderInfo()

On Error GoTo Error_GetProviderInfo

Dim bytSecurityDescriptor() As Byte
Dim lngAlgImpl As Long
Dim lngDataLen As Long
Dim lngKeyIncr As Long
Dim lngRet As Long
Dim lngVersion As Long
Dim oAlgorithm As Node
Dim oAlgorithms As Node
Dim oContainer As Node
Dim oContainers As Node
Dim oGeneral As Node
Dim strContainerName As String
Dim strProviderName As String
Dim strUniqueProviderName As String
Dim udtAlgEnum As PROV_ENUMALGS_EX
Dim X As Long

Screen.MousePointer = vbHourglass

ResetForm

'  Get a handle to the given provider.
mlngHProvider = GetProviderHandle(cboProviders.Text, cboProviders.ItemData(cboProviders.ListIndex), _
                                  txtContainerName.Text)

If mlngHProvider <> 0 Then
    '  Let's get general provider information first.
    Set oGeneral = trvProviders.Nodes.Add(, , , "General Provider Information")
    '  First, get the provider na
    lngRet = CryptGetProvParam(mlngHProvider, PP_NAME, ByVal vbNullString, lngDataLen, 0)
    strProviderName = Space(lngDataLen)
    lngRet = CryptGetProvParam(mlngHProvider, PP_NAME, ByVal strProviderName, lngDataLen, 0)
    If lngRet <> 0 Then
        strProviderName = Left$(strProviderName, lngDataLen - 1)
        trvProviders.Nodes.Add oGeneral, tvwChild, , "Provider Name:  " & strProviderName
    End If
    '  Now get the version number.
    lngDataLen = LenB(lngVersion)
    lngRet = CryptGetProvParam(mlngHProvider, PP_VERSION, lngVersion, lngDataLen, 0)
    If lngRet <> 0 Then
        trvProviders.Nodes.Add oGeneral, tvwChild, , "Version:  " & CStr(lngVersion)
    End If
    '  Find out how the algorithm is implemented.
    lngDataLen = LenB(lngAlgImpl)
    lngRet = CryptGetProvParam(mlngHProvider, PP_IMPTYPE, lngAlgImpl, lngDataLen, 0)
    If lngRet <> 0 Then
        Select Case lngAlgImpl
            Case CRYPT_IMPL_HARDWARE
                trvProviders.Nodes.Add oGeneral, tvwChild, , "Hardware"
            Case CRYPT_IMPL_MIXED
                trvProviders.Nodes.Add oGeneral, tvwChild, , "Mixed"
            Case CRYPT_IMPL_SOFTWARE
                trvProviders.Nodes.Add oGeneral, tvwChild, , "Software"
            Case CRYPT_IMPL_UNKNOWN
                trvProviders.Nodes.Add oGeneral, tvwChild, , "Unknown"
        End Select
    End If
    '  Now find out if this provider has
    '  a hardware RNG implementation.
    '  Note how the arguments are handled in this case.
    lngRet = CryptGetProvParam(mlngHProvider, PP_USE_HARDWARE_RNG, 0&, 0, 0)
    If lngRet = 0 Then
        '  It doesn't.
        trvProviders.Nodes.Add oGeneral, tvwChild, , "No RNG Implementation in Hardware"
    Else
        '  It does.
        trvProviders.Nodes.Add oGeneral, tvwChild, , "RNG Implementation in Hardware is Available"
    End If
    '  Now get the current key container name in use.
    '  Technically, we know what this is, but
    '  it shows how to get it with CryptGetProvParam
    lngDataLen = 0
    lngRet = CryptGetProvParam(mlngHProvider, PP_CONTAINER, ByVal vbNullString, lngDataLen, 0)
    strContainerName = Space(lngDataLen)
    lngRet = CryptGetProvParam(mlngHProvider, PP_CONTAINER, ByVal strContainerName, lngDataLen, 0)
    If lngRet <> 0 Then
        strContainerName = Left$(strContainerName, lngDataLen - 1)
        trvProviders.Nodes.Add oGeneral, tvwChild, , "Current Key Container Name:  " & strContainerName
    End If
    '  Now get the unique container key na
    lngDataLen = 0
    lngRet = CryptGetProvParam(mlngHProvider, PP_UNIQUE_CONTAINER, ByVal vbNullString, lngDataLen, 0)
    strUniqueProviderName = Space(lngDataLen)
    lngRet = CryptGetProvParam(mlngHProvider, PP_UNIQUE_CONTAINER, ByVal strUniqueProviderName, lngDataLen, 0)
    If lngRet <> 0 Then
        strUniqueProviderName = Left$(strUniqueProviderName, lngDataLen - 1)
        trvProviders.Nodes.Add oGeneral, tvwChild, , "Unique Key Container Name:  " & strUniqueProviderName
    End If
    '  Now get the security descriptor for this key set.
    lngDataLen = 0
    lngRet = CryptGetProvParam(mlngHProvider, PP_KEYSET_SEC_DESCR, 0&, lngDataLen, OWNER_SECURITY_INFORMATION)
    ReDim bytSecurityDescriptor(1 To lngDataLen) As Byte
    lngRet = CryptGetProvParam(mlngHProvider, PP_KEYSET_SEC_DESCR, bytSecurityDescriptor(1), lngDataLen, OWNER_SECURITY_INFORMATION)
    If lngRet <> 0 Then
    
    End If
    Set oContainers = trvProviders.Nodes.Add(oGeneral, tvwChild, , "Key Container Names")
    '  Now get all of the container names.
    X = 0
    lngDataLen = MAXUIDLEN
    strContainerName = Space(MAXUIDLEN)
    lngRet = CryptGetProvParam(mlngHProvider, PP_ENUMCONTAINERS, ByVal strContainerName, lngDataLen, CRYPT_FIRST)
    Do Until (lngRet = 0)
        X = X + 1
        '  Show the container na
        '  Note that in this case, we can't use lngDataLen to figure out how long the name is.
        trvProviders.Nodes.Add oContainers, tvwChild, , Left$(strContainerName, InStr(strContainerName, vbNullChar) - 1)
        lngDataLen = MAXUIDLEN
        strContainerName = Space(MAXUIDLEN)
        lngRet = CryptGetProvParam(mlngHProvider, PP_ENUMCONTAINERS, ByVal strContainerName, lngDataLen, 0)
    Loop
    '  Now get all of the algorithm information.
    Set oAlgorithms = trvProviders.Nodes.Add(, , , "Algorithm Information")
    X = 0
    lngDataLen = LenB(udtAlgEnum)
    lngRet = CryptGetProvParam(mlngHProvider, PP_ENUMALGS_EX, udtAlgEnum, lngDataLen, CRYPT_FIRST)
    Do Until (lngRet = 0)
        X = X + 1
        '  Grab the information out of the UDT
        With udtAlgEnum
            Set oAlgorithm = trvProviders.Nodes.Add(oAlgorithms, tvwChild, , Left$(.szName, .dwNameLen - 1))
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "ID:  " & CStr(.aiAlgid)
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Long Name:  " & Left$(.szLongName, .dwLongNameLen - 1)
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Minimum Key Length:  " & CStr(.dwMinLen)
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Maximum Key Length:  " & CStr(.dwMaxLen)
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Default Key Length:  " & CStr(.dwDefaultLen)
            Select Case GetAlgorithmClass(.aiAlgid)
                Case ALG_CLASS_HASH
                    cboHashAlgs.AddItem Left$(.szLongName, .dwLongNameLen - 1)
                    cboHashAlgs.ItemData(cboHashAlgs.NewIndex) = .aiAlgid
                Case ALG_CLASS_DATA_ENCRYPT, ALG_CLASS_KEY_EXCHANGE, ALG_CLASS_MSG_ENCRYPT, ALG_CLASS_SIGNATURE
                    cboKeyAlgs.AddItem Left$(.szLongName, .dwLongNameLen - 1)
                    cboKeyAlgs.ItemData(cboKeyAlgs.NewIndex) = .aiAlgid
            End Select
            '  While we're here, see if we can get the signature key size
            '  if this algorithm is of class ALG_CLASS_SIGNATURE or ALG_CLASS_KEY_EXCHANGE.
            If (GetAlgorithmClass(.aiAlgid) = ALG_CLASS_SIGNATURE) Then
                lngDataLen = LenB(lngKeyIncr)
                lngRet = CryptGetProvParam(mlngHProvider, PP_SIG_KEYSIZE_INC, lngKeyIncr, lngDataLen, 0)
                If lngRet <> 0 Then
                    '  This is a W2K machine, so get the value.
                    trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Signature Key Increment Size:  " & CStr(lngKeyIncr)
                End If
            ElseIf (GetAlgorithmClass(.aiAlgid) = ALG_CLASS_KEY_EXCHANGE) Then
                lngDataLen = LenB(lngKeyIncr)
                lngRet = CryptGetProvParam(mlngHProvider, PP_KEYX_KEYSIZE_INC, lngKeyIncr, lngDataLen, 0)
                If lngRet <> 0 Then
                    '  This is a W2K machine, so get the value.
                    trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Exchange Key Increment Size:  " & CStr(lngKeyIncr)
                End If
            End If
            trvProviders.Nodes.Add oAlgorithm, tvwChild, , "Number of Protocols Supported:  " & CStr(.dwProtocols)
        End With
        lngDataLen = LenB(udtAlgEnum)
        lngRet = CryptGetProvParam(mlngHProvider, PP_ENUMALGS_EX, udtAlgEnum, lngDataLen, 0)
    Loop
    '  Move the hash and combo boxes
    '  to the first element if one exists.
    If cboHashAlgs.ListCount > 0 Then
        cboHashAlgs.ListIndex = 0
    End If
    
    If cboKeyAlgs.ListCount > 0 Then
        cboKeyAlgs.ListIndex = 0
    End If
    oGeneral.Expanded = True
End If

Screen.MousePointer = vbDefault

Exit Sub

Error_GetProviderInfo:

Screen.MousePointer = vbDefault
End Sub

Private Sub CheckOS()

On Error Resume Next

'  If this isn't a W2K box, don't show
'  the Protecting Data tab.
Dim lngRet As Long
Dim udtos As OSVERSIONINFO

udtos.dwOSVersionInfoSize = Len(udtos)
lngRet = GetVersionEx(udtos)

If udtos.dwPlatformId >= VER_PLATFORM_WIN32_NT And udtos.dwMajorVersion >= 5 Then
    '  This is a W2K machine.
    cmdEncryptFileEasy.Enabled = True
    cmdDecryptFileEasy.Enabled = True
Else
    cmdEncryptFileEasy.Enabled = False
    cmdDecryptFileEasy.Enabled = False
End If
End Sub

Private Sub CipherFile(Cipher As CipherType)

On Error Resume Next

Dim strSource As String
Dim strDest As String

If mlngHKey = 0 Then
    mlngHKey = modCryptoAPI.CreateKey(mlngHProvider, cboKeyAlgs.ItemData(cboKeyAlgs.ListIndex), optSalt(0).value, _
        IIf(chkExportKey.value = vbChecked, True, False), optKeyGen(0).value, _
        modCryptoAPI.CreateHash(mlngHProvider, cboHashAlgs.ItemData(cboHashAlgs.ListIndex), Trim$(txtPreImage.Text)))
End If

If mlngHKey <> 0 Then
    strSource = GetSourceFile(hwnd)
    If strSource <> vbNullString Then
        strDest = GetDestinationFile(hwnd)
        If strDest <> vbNullString Then
            modCryptoAPI.CipherFile Cipher, strSource, strDest, mlngHKey
        End If
    End If
End If
End Sub

Private Sub CipherFileEasy(Cipher As CipherType)

On Error Resume Next

Dim lngRet As Long
Dim strSource As String

strSource = GetSourceFile(hwnd)

If strSource <> vbNullString Then
    If Cipher = Encrypt Then
        lngRet = CryptEncryptFile(strSource)
    Else
        lngRet = CryptDecryptFile(strSource, 0)
    End If
End If
End Sub

Private Sub CreateHash()

On Error GoTo Error_CreateHash

Dim lngDataLen As Long
Dim lngHashSize As Long
Dim lngRet As Long
Dim oHashNode As Node
Dim strHash As String
Dim X As Long

Screen.MousePointer = vbHourglass

trvHashResults.Nodes.Clear

If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

If mlngHProvider <> 0 And cboHashAlgs.ListIndex >= 0 Then
    If Trim$(txtPreImage.Text) <> vbNullString Then
        '  Now create the hash value.
        mlngHHash = modCryptoAPI.CreateHash(mlngHProvider, cboHashAlgs.ItemData(cboHashAlgs.ListIndex), _
            Trim$(txtPreImage.Text))
        If mlngHHash <> 0 Then
            '  Display the hash value.
            DisplayHashValue mlngHHash
            CryptDestroyHash mlngHHash
        Else
            ShowCryptoAPIError
        End If
    End If
End If

Screen.MousePointer = vbDefault

Exit Sub

Error_CreateHash:

Screen.MousePointer = vbDefault
End Sub

Private Sub CreateKey()

On Error GoTo Error_CreateKey

Dim bytIV() As Byte
Dim lngBlockLen As Long
Dim lngBlockMode As Long
Dim lngDataLen As Long
Dim lngFlags As Long
Dim lngKeySize As Long
Dim lngPermissions As Long
Dim lngRet As Long
Dim oKeyBlockNode As Node
Dim oKeyNode As Node

Screen.MousePointer = vbHourglass

If mlngHKey <> 0 Then
    CryptDestroyKey mlngHKey
End If

trvKeyResults.Nodes.Clear

If mlngHProvider <> 0 And cboKeyAlgs.ListIndex >= 0 Then
    mlngHKey = modCryptoAPI.CreateKey(mlngHProvider, cboKeyAlgs.ItemData(cboKeyAlgs.ListIndex), optSalt(0).value, _
        IIf(chkExportKey.value = vbChecked, True, False), optKeyGen(0).value, _
        modCryptoAPI.CreateHash(mlngHProvider, cboHashAlgs.ItemData(cboHashAlgs.ListIndex), Trim$(txtPreImage.Text)))
    If mlngHKey <> 0 Then
        '  Get the key information.
        '  First check to see if we can read key information.
        lngDataLen = LenB(lngPermissions)
        lngRet = CryptGetKeyParam(mlngHKey, KP_PERMISSIONS, lngPermissions, lngDataLen, 0)
        If lngRet <> 0 Then
            If lngPermissions And CRYPT_READ Then
                '  OK, let's get the key size.
                Set oKeyNode = trvKeyResults.Nodes.Add(, , , "Key Results")
                lngDataLen = LenB(lngKeySize)
                lngRet = CryptGetKeyParam(mlngHKey, KP_KEYLEN, lngKeySize, lngDataLen, 0)
                If lngRet <> 0 Then
                    trvKeyResults.Nodes.Add oKeyNode, tvwChild, , "Length:  " & CStr(lngKeySize)
                End If
                '  Now let's see if this is a block or stream cipher.
                lngDataLen = LenB(lngBlockLen)
                lngRet = CryptGetKeyParam(mlngHKey, KP_BLOCKLEN, lngBlockLen, lngDataLen, 0)
                If lngRet <> 0 Then
                    If lngBlockLen = 0 Then
                        '  It's a stream key.
                        trvKeyResults.Nodes.Add oKeyNode, tvwChild, , "Stream Cipher"
                    Else
                        '  We can adjust the IV vector.
                        '  Check to see if we have write permissions.
                        lngDataLen = LenB(lngPermissions)
                        lngRet = CryptGetKeyParam(mlngHKey, KP_PERMISSIONS, lngPermissions, lngDataLen, 0)
                        '  It's a block cipher.
                        Set oKeyBlockNode = trvKeyResults.Nodes.Add(oKeyNode, tvwChild, , "Block Cipher")
                        trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Length:  " & CStr(lngBlockLen)
                        '  OK, let's get the IV vector.
                        '  We know that the array will be of size (lngblocklen/8)
                        ReDim bytIV(0 To (lngBlockLen / 8 - 1)) As Byte
                        lngDataLen = lngBlockLen / 8
                        lngRet = CryptGetKeyParam(mlngHKey, KP_IV, bytIV(0), lngDataLen, 0)
                        '  We shouldn't need to check for ERROR_MORE_DATA here.
                        trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "IV Value:  " & _
                                                GetHexString(bytIV)
                        '  Get the block mode
                        lngDataLen = LenB(lngBlockMode)
                        lngRet = CryptGetKeyParam(mlngHKey, KP_MODE, lngBlockMode, lngDataLen, 0)
                        If lngRet <> 0 Then
                            Select Case lngBlockMode
                                Case CRYPT_MODE_CBC
                                    trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Cipher Block Chaining Mode"
                                Case CRYPT_MODE_CFB
                                    trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Cipher Feedback Mode"
                                Case CRYPT_MODE_ECB
                                    trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Electronic Codebook"
                                Case CRYPT_MODE_OFB
                                    trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Output Feedback Mode"
                                Case Else
                                    trvKeyResults.Nodes.Add oKeyBlockNode, tvwChild, , "Unknown Block Mode"
                            End Select
                        End If
                    End If
                End If
                '  Now display the key values.
                DisplayKeyInfo mlngHKey, trvKeyResults.Nodes.Add(oKeyNode, tvwChild, , "Key Information")
            Else
                Set oKeyNode = trvKeyResults.Nodes.Add(, , , "Cannot Read Key Values")
            End If
            oKeyNode.Expanded = True
        Else
            ShowCryptoAPIError
        End If
    CryptDestroyKey mlngHKey
    Else
        ShowCryptoAPIError
    End If
End If

Screen.MousePointer = vbDefault

Exit Sub

Error_CreateKey:

If mlngHKey <> 0 Then
    CryptDestroyKey mlngHKey
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub DisplayHashValue(HashHandle As Long)

On Error GoTo Error_DisplayHashValue

Dim bytHash() As Byte
Dim lngDataLen As Long
Dim lngHashSize As Long
Dim lngRet As Long
Dim oHashNode As Node

'  Get the hash value.
Set oHashNode = trvHashResults.Nodes.Add(, , , "Hash Results")
lngDataLen = LenB(lngHashSize)
lngRet = CryptGetHashParam(HashHandle, HP_HASHSIZE, lngHashSize, lngDataLen, 0&)
trvHashResults.Nodes.Add oHashNode, tvwChild, , "Hash Size:  " & CStr(lngHashSize)

ReDim bytHash(0 To (lngHashSize - 1)) As Byte
lngDataLen = lngHashSize
lngRet = CryptGetHashParam(HashHandle, HP_HASHVAL, bytHash(0), lngDataLen, 0&)
trvHashResults.Nodes.Add oHashNode, tvwChild, , "Hash Value:  " _
                         & GetHexString(bytHash)
oHashNode.Expanded = True

Exit Sub

Error_DisplayHashValue:
End Sub

Private Sub DisplayKeyInfo(KeyHandle As Long, BaseTreeNode As Node)

On Error GoTo Error_DisplayKeyInfo

Dim bytKeyValue() As Byte
Dim bytKey() As Byte
Dim lngAlgID As Long
Dim lngBitCount As Long
Dim lngDataLen As Long
Dim lngExponent As Long
Dim lngMagic As Long
Dim lngRet As Long

lngRet = CryptExportKey(KeyHandle, 0, PUBLICKEYBLOB, 0, 0, lngDataLen)
'  Now re-invoke with the correct buffer.
ReDim bytKey(0 To (lngDataLen - 1)) As Byte
lngRet = CryptExportKey(KeyHandle, 0, PUBLICKEYBLOB, 0, bytKey(0), lngDataLen)

If lngRet <> 0 Then
    '  The first byte just defines the blob type; we know
    '  this, so simply display it.
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Public Key Blob"
    '  Get the version
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Version:  " & CStr(bytKey(1))
    '  The reserved value can be ignored.
    '  Now show the algID.  Note that ALG_ID is an unsigned int.
    CopyMemory lngAlgID, bytKey(4), 4
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Algorithm ID:  " & CStr(lngAlgID)
    '  Show the "magic" number.
    CopyMemory lngMagic, bytKey(8), 4
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Magic Value:  " & CStr(lngMagic)
    '  Show the number of bits in the modulus.
    CopyMemory lngBitCount, bytKey(12), 4
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Bit Count:  " & CStr(lngBitCount)
    '  Show the exponent value.
    CopyMemory lngExponent, bytKey(16), 4
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Exponent Value:  " & CStr(lngExponent)
    '  Now display the actual key value.
    ReDim bytKeyValue(0 To ((lngBitCount / 8) - 1)) As Byte
    CopyMemory bytKeyValue(0), bytKey(20), (lngBitCount / 8)
    trvKeyResults.Nodes.Add BaseTreeNode, tvwChild, , "Key Value:  " & GetHexString(bytKeyValue)
End If

Exit Sub

Error_DisplayKeyInfo:
End Sub

Private Function GetHexString(BaseBytes() As Byte)

Dim X As Long

For X = LBound(BaseBytes) To UBound(BaseBytes)
    If Len(Hex$(BaseBytes(X))) = 2 Then
        GetHexString = GetHexString & Hex$(BaseBytes(X)) & " "
    Else
        GetHexString = GetHexString & "0" & Hex$(BaseBytes(X)) & " "
    End If
Next X

GetHexString = Trim$(GetHexString)
End Function

Private Sub LoadProviders()

Dim oProvider As clsProvider
Dim oProviders As clsProviders
Dim X As Long

Set oProviders = modCryptoAPI.GetProviders

If Not oProviders Is Nothing Then
    For Each oProvider In oProviders
        cboProviders.AddItem oProvider.Name
        cboProviders.ItemData(cboProviders.NewIndex) = oProvider.ProviderType
    Next oProvider
    cboProviders.ListIndex = 0
End If
End Sub

Private Sub ResetForm()

On Error Resume Next

'  Reset provider tab except for the provider list.
trvProviders.Nodes.Clear

'  Reset the hash information
txtPreImage.Text = vbNullString
trvHashResults.Nodes.Clear
cboHashAlgs.Clear

'  Reset the key information
trvKeyResults.Nodes.Clear
cboKeyAlgs.Clear
End Sub

Private Sub ShowCipherUI()

On Error Resume Next

cmdEncryptFile.Enabled = True
cmdDecryptFile.Enabled = True

If optKeyGen(0).value = False And Trim$(txtPreImage.Text) = vbNullString Then
    cmdEncryptFile.Enabled = False
    cmdDecryptFile.Enabled = False
End If
End Sub

Private Sub ShowHashUI()

If mlngHProvider = 0 Then
    '  Can't create a hash value just yet.
    cmdCreateHash.Enabled = False
    cmdSignAndVerify.Enabled = False
    cmdSignDistortVerify.Enabled = False
Else
    cmdCreateHash.Enabled = True
    cmdSignAndVerify.Enabled = True
    cmdSignDistortVerify.Enabled = True
End If
End Sub

Private Sub ShowKeyUI()

If Trim$(txtPreImage.Text) = vbNullString Then
    '  The user can't create a key
    '  based off of hash data.
    optKeyGen(1).Enabled = False
Else
    optKeyGen(1).Enabled = True
End If

If mlngHProvider = 0 Then
    '  Can't create a key value just yet.
    cmdCreateKey.Enabled = False
Else
    cmdCreateKey.Enabled = True
End If
End Sub

Private Sub SignAndVerify()

On Error GoTo Error_SignAndVerify

Dim bytSig() As Byte
Dim lngAlgID As Long
Dim lngHUserKey As Long
Dim lngRet As Long
Dim lngRetLen As Long
Dim oNode As Node

Screen.MousePointer = vbHourglass

If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

trvHashResults.Nodes.Clear

'  First, create a new hash.
lngAlgID = cboHashAlgs.ItemData(cboHashAlgs.ListIndex)
mlngHHash = modCryptoAPI.CreateHash(mlngHProvider, lngAlgID, _
    Trim$(txtPreImage.Text))
    
DisplayHashValue mlngHHash

If mlngHHash <> 0 Then
    '  OK, now sign it.  We'll use AT_SIGNATURE for the private key.
    lngRetLen = 0
    lngRet = CryptSignHash(mlngHHash, AT_SIGNATURE, vbNullString, 0, ByVal 0&, lngRetLen)
    '  Re-invoke with the correct signature size.
    ReDim bytSig(0 To (lngRetLen - 1)) As Byte
    lngRet = CryptSignHash(mlngHHash, AT_SIGNATURE, vbNullString, 0, bytSig(0), lngRetLen)
    If lngRet <> 0 Then
        '  Successful signing.
        '  Now verify it (note that we have to recreate the hash).
        CryptDestroyHash mlngHHash
        mlngHHash = modCryptoAPI.CreateHash(mlngHProvider, lngAlgID, _
            Trim$(txtPreImage.Text))
        If mlngHHash <> 0 Then
            lngRet = CryptGetUserKey(mlngHProvider, AT_SIGNATURE, lngHUserKey)
            lngRet = CryptVerifySignature(mlngHHash, bytSig(0), lngRetLen, lngHUserKey, vbNullString, 0)
            If lngRet <> 0 Then
                '  Signature was verified.
                trvHashResults.Nodes.Add , , , "Signature " & GetHexString(bytSig) & " was verified."
            Else
                trvHashResults.Nodes.Add , , , "Signature was incorrect - error code:  " & Err.LastDllError & "."
            End If
        End If
    End If
End If

'  Destroy the hash value.
If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

Screen.MousePointer = vbDefault

Exit Sub

Error_SignAndVerify:

If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub SignDistortAndVerify()

On Error GoTo Error_SignDistoryAndVerify

Dim bytSig() As Byte
Dim lngAlgID As Long
Dim lngHUserKey As Long
Dim lngRet As Long
Dim lngRetLen As Long
Dim oNode As Node

Screen.MousePointer = vbHourglass

If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

trvHashResults.Nodes.Clear

'  First, create a new hash.
lngAlgID = cboHashAlgs.ItemData(cboHashAlgs.ListIndex)
mlngHHash = modCryptoAPI.CreateHash(mlngHProvider, lngAlgID, _
    Trim$(txtPreImage.Text))

DisplayHashValue mlngHHash

If mlngHHash <> 0 Then
    '  OK, now sign it.  We'll use AT_SIGNATURE for the private key.
    lngRetLen = 0
    lngRet = CryptSignHash(mlngHHash, AT_SIGNATURE, vbNullString, 0, ByVal 0&, lngRetLen)
    '  Re-invoke with the correct signature size.
    ReDim bytSig(0 To (lngRetLen - 1)) As Byte
    lngRet = CryptSignHash(mlngHHash, AT_SIGNATURE, vbNullString, 0, bytSig(0), lngRetLen)
    If lngRet <> 0 Then
        '  Successful signing.
        '  Now distort it.
        bytSig(LBound(bytSig)) = 40
        bytSig(UBound(bytSig)) = 40
        '  Now verify it (note that we have to recreate the hash).
        CryptDestroyHash mlngHHash
        mlngHHash = modCryptoAPI.CreateHash(mlngHProvider, lngAlgID, _
            Trim$(txtPreImage.Text))
        If mlngHHash <> 0 Then
            lngRet = CryptGetUserKey(mlngHProvider, AT_SIGNATURE, lngHUserKey)
            lngRet = CryptVerifySignature(mlngHHash, bytSig(0), lngRetLen, lngHUserKey, vbNullString, 0)
            If lngRet <> 0 Then
                '  Signature was verified.
                trvHashResults.Nodes.Add , , , "Signature " & GetHexString(bytSig) & " was verified."
            Else
                trvHashResults.Nodes.Add , , , "Signature was incorrect - error code:  " & Err.LastDllError & "."
            End If
        End If
    End If
End If

'  Destroy the hash value.
If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

Screen.MousePointer = vbDefault

Exit Sub

Error_SignDistoryAndVerify:

If mlngHHash <> 0 Then
    CryptDestroyHash mlngHHash
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub txtPassCert_Change()
    INI.DeleteKey "SETTING", "CERTIFICATE_PASSWORD"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_PASSWORD", txtPassCert.Text
End Sub

Private Sub txtValidFrom_Change()
    INI.DeleteKey "SETTING", "CERTIFICATE_VALID_FROM"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_VALID_FROM", txtValidFrom.Text
End Sub


Private Sub txtValidTo_Change()
    INI.DeleteKey "SETTING", "CERTIFICATE_VALID_TO"
    INI.CreateKeyValue "SETTING", "CERTIFICATE_VALID_TO", txtValidTo.Text
End Sub



Private Sub CreateXML(strFileName As String)
    Dim intFile As Integer
    On Local Error GoTo ErrorXML
    intFile = FreeFile()
    Open strFileName For Output As #intFile
        Print #intFile, "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "utf-8" & """" & " ?>"
        Print #intFile, "<wap-provisioningdoc>"
        Print #intFile, "<characteristic type=" & """" & "CertificateStore" & """" & ">"
        Print #intFile, Tab(10); "<characteristic type=" & """" & "Privileged Execution Trust Authorities" & """" & ">"
        Print #intFile, Tab(15); "<characteristic type=" & """"; txtSHA1.Text; """" & ">"
        Print #intFile, Tab(20); "<parm name=" & """" & "EncodedCertificate" & """" & " value=" & """"; txtBase64.Text; """" & " />"
        Print #intFile, Tab(15); "</characteristic>"
        Print #intFile, Tab(10); "</characteristic>"
        Print #intFile, "</characteristic>"
        Print #intFile, "<characteristic type=" & """" & "CertificateStore" & """" & ">"
        Print #intFile, Tab(10); "<characteristic type=" & """" & "SPC" & """" & ">"
        Print #intFile, Tab(15); "<characteristic type=" & """"; txtSHA1.Text; """" & ">"
        Print #intFile, Tab(20); "<parm name=" & """" & "EncodedCertificate" & """" & " value=" & """"; txtBase64.Text; """" & " />"
        Print #intFile, Tab(20); "<parm name=" & """" & "Role" & """" & " value=" & """" & "222" & """" & " />"
        Print #intFile, Tab(15); "</characteristic>"
        Print #intFile, Tab(10); "</characteristic>"
        Print #intFile, "</characteristic>"
        Print #intFile, "</wap-provisioningdoc>"
    Close #intFile
Exit Sub
ErrorXML:
    Err.Clear
End Sub

Private Function DecryptMessage(ByRef oMsg As CDO.Message) As Boolean
    Dim oDecryptedMsg As New CDO.Message
    Dim oStream As New ADODB.stream
    Dim iDsrc As IDataSource
    Dim oEnvelopedData As New CAPICOM.EnvelopedData
    Dim byteDecryptedMessage() As Byte
         
    On Error GoTo ErrorHandler
    
    ' .... Decrypt content
    Call oEnvelopedData.Decrypt(oMsg.BodyPart.GetEncodedContentStream.ReadText)
    
    ' .... Convert the message to a byte array
    byteDecryptedMessage = oEnvelopedData.Content

    ' .... Load the decrypted message into a stream
    oStream.Open
    oStream.Type = adTypeBinary
    oStream.Write byteDecryptedMessage

    Set iDsrc = oDecryptedMsg
    iDsrc.OpenObject oStream, "_Stream"
    
    ' .... Return the status
    Set oMsg = oDecryptedMsg
    DecryptMessage = True

GoTo CleanUp

ErrorHandler:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title
    
    ' .... Return the false values
    DecryptMessage = False
    Set oMsg = Nothing
    
CleanUp:
    ' .... Clean up
    Set oEnvelopedData = Nothing
    Set oDecryptedMsg = Nothing
    Set oStream = Nothing
    Set iDsrc = Nothing
End Function

Private Function EncryptMessage(ByRef oMsg As CDO.Message, oRecipients As Certificates) As Boolean
'******************************************************************************
'
' Function:     EncryptMessage
'
' Parameters:   oMsg        -   A CDO object representing a properly formed MIME
'                               message. [in/out]
'
'               oRecipients  -  A collection of CAPICOM certificate objects in
'                               which should be capable of decrypting this
'                               message. [in]
'
'******************************************************************************
    Dim oEncryptedMsg As New CDO.Message
    Dim oBodyPart As CDO.IBodyPart
    Dim cFields As ADODB.Fields
    Dim oStream As ADODB.stream
    Dim oEnvelopedData As New CAPICOM.EnvelopedData
    Dim oRecipient As CAPICOM.Certificate
    Dim szEncMessage, byteEncMessage() As Byte
     
    ' .... Copy input into output message
    oEncryptedMsg.DataSource.OpenObject oMsg, cdoIMessage
    
    ' .... Set up main bodypart
    Set oBodyPart = oEncryptedMsg.BodyPart
    oBodyPart.ContentMediaType = "application/pkcs7-mime;" & vbCrLf & "smime-type=enveloped-data;" & vbCrLf & "name=smime.p7m;"
    oBodyPart.ContentTransferEncoding = "base64"
    oBodyPart.Fields("urn:schemas:mailheader:content-disposition") = "attachment;FileName=""smime.p7m"""
    oBodyPart.Fields("urn:schemas:mailheader:date").value = oMsg.Fields("urn:schemas:mailheader:date").value
    oBodyPart.Fields.Update
    
    ' .... Add each of the passed in recipients to the EnvelopedData recipient's collection
    For Each oRecipient In oRecipients
        oEnvelopedData.Recipients.Add oRecipient
    Next
    
    ' .... Encrypt content
    oEnvelopedData.Content = StrConv(oMsg.BodyPart.GetStream.ReadText, vbFromUnicode)
    szEncMessage = oEnvelopedData.Encrypt(CAPICOM_ENCODE_BINARY)
    
    ' .... Get the string data as a byte array
    byteEncMessage = szEncMessage
    
    ' .... Write the CMS blob into the main bodypart and let CDO do the base64 encoding
    Set oStream = oEncryptedMsg.BodyPart.GetDecodedContentStream
    oStream.Type = adTypeBinary
    oStream.Write byteEncMessage
    oStream.Flush
       
    ' .... Return out finished message
    EncryptMessage = True
    Set oMsg = oEncryptedMsg

GoTo CleanUp

ErrorHandler:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title
    Err.Clear
    EncryptMessage = False
    Set oMsg = Nothing

CleanUp:
    ' .... Clean up
    Set oBodyPart = Nothing
    Set oEnvelopedData = Nothing
    Set oStream = Nothing
    Set oRecipient = Nothing
    Set oEncryptedMsg = Nothing
    Set oBodyPart = Nothing
    Set cFields = Nothing
End Function

Private Function FindRecipientByEmail(ByVal szEmail, ByRef oRecipient As CAPICOM.Certificate) As Boolean
    Dim oStore As New CAPICOM.Store
    Dim oCertificates As New CAPICOM.Certificates
    Dim oCertificate As New CAPICOM.Certificate
    
    On Error GoTo ErrorHandler
    
    ' .... Open the AddressBook store to see if we can find their certificate in their
    oStore.Open CAPICOM_CURRENT_USER_STORE, cmbsStoresList.List(cmbsStoresList.ListIndex), CAPICOM_STORE_OPEN_READ_ONLY
    
    ' .... We are only interested in those certificates that are explicitly good for secure email and key encipherment
    Set oCertificates = oStore.Certificates.Find(CAPICOM_CERTIFICATE_FIND_APPLICATION_POLICY, "Secure Email").Find(CAPICOM_CERTIFICATE_FIND_KEY_USAGE, "KeyEncipherment", True)
  
    ' .... This simply picks the first match on email address
    For Each oCertificate In oCertificates
        If InStr(1, szEmail, "@") Then
            ' .... Looks like a complete email address
            If (oCertificate.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME) = szEmail) Then
                    FindRecipientByEmail = True
                Exit For
            End If
        Else
            ' .... Looks like a partial email address or alias
            If (InStr(1, oCertificate.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME), szEmail)) Then
                FindRecipientByEmail = True
                Exit For
            End If
        End If
    Next

If (FindRecipientByEmail = True) Then
    Set oRecipient = oCertificate
Else
    MsgBox "Unable to find Encryption certificate for '" & szEmail & "', this recipient will not be included.", vbExclamation, App.Title
    Set oRecipient = Nothing
End If

GoTo CleanUp

ErrorHandler:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title

CleanUp:
Set oStore = Nothing
Set oCertificates = Nothing
Set oCertificate = Nothing
End Function

Private Function GetCertForSignature(subject As String) As CAPICOM.Certificate
    Dim cert As CAPICOM.Certificate
    Dim st As New CAPICOM.Store
    st.Open CAPICOM_CURRENT_USER_STORE, cmbsStoresList.List(cmbsStoresList.ListIndex), CAPICOM_STORE_OPEN_READ_ONLY
    For Each cert In st.Certificates
        If (cert.IsValid) And _
           (StrComp(cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME + CAPICOM_CERT_INFO_ISSUER_EMAIL_NAME + _
           CAPICOM_CERT_INFO_ISSUER_SIMPLE_NAME + CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME + _
           CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME), subject, vbTextCompare) = 0) And _
           (cert.KeyUsage.IsDigitalSignatureEnabled) Then
                Set GetCertForSignature = cert
            Exit Function
        End If
    Next
    Set GetCertForSignature = Nothing
End Function

Private Function GetContent(oInMsg As CDO.Message) As String
    Dim iStart As Integer, iLength As Integer
    Dim szMessage, szBodyPart
    szMessage = oInMsg.GetStream.ReadText
    szBodyPart = "--" + oInMsg.BodyPart.GetFieldParameter("urn:schemas:mailheader:content-type", "boundary") + vbCrLf
    iStart = InStr(1, szMessage, szBodyPart) + Len(szBodyPart)
    iLength = InStr((iStart + 1), szMessage, szBodyPart) - iStart - 2
    GetContent = Mid(szMessage, iStart, iLength)
End Function

Private Function GetLoggedInUser(sUserName As String) As Boolean
    Dim sBuff As String * 25
    Dim lRet As Long
    GetLoggedInUser = True

    ' .... Get the user name, remove NULLs, and trim trailing spaces
    lRet = GetUserName(sBuff, 25)
    sUserName = Trim$(Left(sBuff, InStr(sBuff, Chr(0)) - 1))

    ' .... Return false if no name is returned
    If sUserName = vbNullString Then GetLoggedInUser = False
End Function

Private Function GetSignature(oInMsg As CDO.Message) As String
    If InStr(1, oInMsg.Fields.Item("urn:schemas:mailheader:content-disposition").value, "attachment", vbTextCompare) <> 0 Then
        GetSignature = oInMsg.BodyPart.GetEncodedContentStream.ReadText
    Else
        GetSignature = oInMsg.BodyPart.BodyParts(2).GetEncodedContentStream.ReadText
    End If
End Function

Private Function IsEncrypted(oInMsg As CDO.Message) As Boolean
    Dim szContentType As String
    szContentType = oInMsg.BodyPart.Fields.Item("urn:schemas:mailheader:content-type").value
    If (InStr(1, szContentType, "enveloped-data", vbTextCompare) <> 0) Then
        IsEncrypted = True
    Else
        IsEncrypted = False
    End If
End Function

Private Function IsSigned(oInMsg As CDO.Message) As Boolean
    Dim szContentType As String
    szContentType = oInMsg.BodyPart.Fields.Item("urn:schemas:mailheader:content-type").value
    If ((InStr(1, szContentType, "application/x-pkcs7-signature", vbTextCompare) <> 0) Or (InStr(1, szContentType, "signed-data", vbTextCompare) <> 0)) Then
        IsSigned = True
    Else
        IsSigned = False
    End If
End Function

Private Function LoadMessage(FileName As String) As Boolean
    Dim oStream As New ADODB.stream
    Dim iDsrc As IDataSource
    Dim oAttribute As CAPICOM.Attribute
    Dim szSignature As String
    Dim oUtilities As New CAPICOM.Utilities
    
    On Error GoTo ErrorHandler
    
    ' .... Load the message from disk
    oStream.Open
    oStream.LoadFromFile FileName

    Set iDsrc = oMessage
    iDsrc.OpenObject oStream, "_Stream"
    oStream.Close
        
    ' .... Hide the encryption status buttons if the message is not encrypted
    If (IsEncrypted(oMessage) = False) Then
        btnBadEnc.Visible = False
        btnGoodEnc.Visible = False
    Else
        ' .... If the message is encrypted then update the form accordingly
        If (DecryptMessage(oMessage)) Then
            btnGoodEnc.Visible = True
        Else
            btnBadEnc.Visible = True
        End If
    End If
    
    ' .... Load the message contents into the form
    txtFromValue.Text = oMessage.From
    txtToValue.Text = oMessage.To
    txtCCValue.Text = oMessage.CC
    txtSubjectValue.Text = oMessage.subject
    txtDateValue.Text = oUtilities.UTCTimeToLocalTime(oMessage.Fields("urn:schemas:mailheader:date").value)

    ' .... Lock the fields that could potentialy be modified since we are looking at an existing mail
    txtMessageBody.Locked = True
    txtMessageBody.Text = oMessage.TextBody
    
    ' .... Hide the signature status buttons if the message is not signed
    If (IsSigned(oMessage) = False) Then
            btnGoodSig.Visible = False
            btnBadSig.Visible = False
    End If

    ' .... If the message is signed then update the form accordingly
    If (IsSigned(oMessage) = True And (IsEncrypted(oMessage) = False)) Then

        ' .... We must verify the message to get the signing certificate
        If (VerifyMessage(oMessage) = True) Then
            txtFromValue.Text = oSigner.Certificate.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME)
        
            For Each oAttribute In oSigner.AuthenticatedAttributes
                If (oAttribute.Name = CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME) Then
                    txtDateValue.Text = oUtilities.UTCTimeToLocalTime(oAttribute.value)
                End If
            Next
            
            btnGoodSig.Visible = True
        Else
            btnBadSig.Visible = True
        End If
    
    End If
GoTo CleanUp
ErrorHandler:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title

CleanUp:
    Set oStream = Nothing
    Set iDsrc = Nothing
    Set oAttribute = Nothing
    Set oUtilities = Nothing
End Function

Private Function ResolveNames(ByRef szNames As String) As CAPICOM.Certificates
    Dim oRecipients As New CAPICOM.Certificates
    Dim oRecipient As New CAPICOM.Certificate
    Dim aNames As Variant, vName As Variant
    Dim bFound As Boolean
    
    On Error GoTo ErrorHandler
    
    ' .... Normalize the delimiter to be ;
    szNames = Replace(szNames, ",", ";")

    ' .... Convert the ; delimited list to an array
    aNames = Split(szNames, ";")
    szNames = ""
    
    For Each vName In aNames
        If (FindRecipientByEmail(vName, oRecipient)) Then
            oRecipients.Add oRecipient
            szNames = szNames + oRecipient.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME) + ";"
        End If
    Next
    
    ' .... Trim the trailing ; delimiter
    If Len(szNames) > 1 Then
        szNames = Mid(szNames, 1, Len(szNames) - 1)
    End If
    
GoTo CleanUp

ErrorHandler:
        MsgBox Err.Number & ": " & Err.Description, vbExclamation, App.Title
    Set ResolveNames = Nothing
    
CleanUp:
    ' .... Clean up
    Set oRecipient = Nothing
    
    ' .... Return Recipient collection
    Set ResolveNames = oRecipients
End Function

Private Function SignMessage(ByRef oMsg As CDO.Message, bClear As Boolean) As Boolean
'******************************************************************************
'
' Function:     SignMessage
'
' Parameters:   oMsg    -   A CDO object representing a properly formed MIME
'                           message. [in/out]
'
'               bClear  -   a boolean specifying if the message is to be signed
'                           using a detached PKCS7 or attached PKCS7. [in]
'
'
' Purpose:      Return a S/MIME message derived from the passed in message
'
'******************************************************************************
    Dim oSignedMsg As New CDO.Message
    Dim oBodyPart As CDO.IBodyPart
    Dim cFields As ADODB.Fields
    Dim oStream As ADODB.stream
    Dim oSignedData As New CAPICOM.SignedData
    Dim oUtilities As New CAPICOM.Utilities
    Dim oAttribute As New CAPICOM.Attribute
    Dim oSignerCertificate As CAPICOM.Certificate
    Dim cSignerCertificates As CAPICOM.Certificates
    Dim oStore As New CAPICOM.Store
    Dim szSignature, byteSignature() As Byte

    On Error GoTo ErrorHandler
    
    ' .... Create the SignedData object we will use to create the PKCS7
    Set oSignedData = New CAPICOM.SignedData
    
    ' .... Create the new message
    Set oSignedMsg = New CDO.Message
    
    ' .... Select the signer certificate
    oStore.Open CAPICOM_CURRENT_USER_STORE, "My, CAPICOM_STORE_OPEN_READ_ONLY"
    Set cSignerCertificates = oStore.Certificates.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CERT_KEY_SPEC_PROP_ID).Find(CAPICOM_CERTIFICATE_FIND_APPLICATION_POLICY, "Secure Email")

    Select Case cSignerCertificates.count
        Case 0
            MsgBox "Error: No signing certificate can be found! Please change your Certificate Sores!", vbExclamation, App.Title
        Case 1
            oSigner.Certificate = cSignerCertificates(1)
        Case Else
            Set cSignerCertificates = cSignerCertificates.Select("S/MIME Certificates", "Please select a certificate to sign with.")
            If (cSignerCertificates.count = 0) Then
                    MsgBox "Error: Certificate selection dialog was cancelled.", vbExclamation, App.Title
                Exit Function
            End If
            oSigner.Certificate = cSignerCertificates(1)
    End Select
    
    ' .... Set the from field based off of the selected certificate
    oSignedMsg.From = oSigner.Certificate.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME)

        
    ' .... Set the signing time in UTC time
    Set oAttribute = New CAPICOM.Attribute
    oAttribute.Name = CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME
    oAttribute.value = oUtilities.LocalTimeToUTCTime(Now)
    oSigner.AuthenticatedAttributes.Add oAttribute
    
    Select Case bClear
    Case True
        ' .... This is to be a clear text signed message so we need to copy the interesting
        ' .... parts (sender, recipient, and subject) into the new header
        oSignedMsg.To = oMsg.To
        oSignedMsg.CC = oMsg.CC
        oSignedMsg.subject = oMsg.subject

        Set oBodyPart = oSignedMsg.BodyPart.AddBodyPart
        Set cFields = oBodyPart.Fields
        cFields.Item(cdoContentType).value = oMsg.BodyPart.BodyParts(1).Fields.Item(cdoContentType).value
        cFields.Update
        
        Set oStream = oBodyPart.GetDecodedContentStream
        oStream.WriteText oMsg.BodyPart.BodyParts(1).GetDecodedContentStream.ReadText
        oStream.Flush
                        
        ' .... Set the content to be signed
        oSignedData.Content = StrConv(oSignedMsg.BodyPart.BodyParts(1).GetStream.ReadText, vbFromUnicode)
                
        ' .... Sign the content
        szSignature = oSignedData.Sign(oSigner, True, CAPICOM_ENCODE_BINARY)
        
        ' .... Get the string data as a byte array
        byteSignature = szSignature
        
        ' .... Attach the signature and let CDO base64 encode it
        Set oBodyPart = oSignedMsg.BodyPart.AddBodyPart
        Set cFields = oBodyPart.Fields
        oBodyPart.Fields.Item("urn:schemas:mailheader:content-type").value = "application/x-pkcs7-signature" & vbCrLf & "Name = ""smime.p7s"""
        oBodyPart.Fields.Item("urn:schemas:mailheader:content-transfer-encoding").value = "base64"
        oBodyPart.Fields.Item("urn:schemas:mailheader:content-disposition").value = "attachment;" & vbCrLf & "FileName=""smime.p7s"""
        cFields.Update
        
        Set oStream = oBodyPart.GetDecodedContentStream
        oStream.Type = ADODB.StreamTypeEnum.adTypeBinary
        oStream.Write (byteSignature)
        oStream.Flush
        
        ' .... Set the messages content type, this needs to be done last to ensure it is not changed when we add the BodyParts
        oSignedMsg.Fields.Item("urn:schemas:mailheader:content-type").value = "multipart/signed;" & vbCrLf & "protocol=""application/x-pkcs7-signature"";" & vbCrLf & "micalg=SHA1;"
        oSignedMsg.Fields.Update
        
    Case False
        ' .... This is to be a opaquely signed message so we need to copy the entire message into our
        ' .... new encrypted message
        oSignedMsg.DataSource.OpenObject oMsg, cdoIMessage
        
        ' .... Set up main bodypart
        Set oBodyPart = oSignedMsg.BodyPart
        oBodyPart.ContentMediaType = "application/pkcs7-mime;" & vbCrLf & "smime-type=signed-data;" & vbCrLf & "name=""smime.p7m"""
        oBodyPart.ContentTransferEncoding = "base64"
        oBodyPart.Fields("urn:schemas:mailheader:content-disposition") = "attachment;" & vbCrLf & "FileName=""smime.p7m"""
        oBodyPart.Fields.Update
          
        ' .... Set the from field based off of the selected certificate
        oMsg.From = oSigner.Certificate.GetInfo(CAPICOM_CERT_INFO_SUBJECT_EMAIL_NAME)
        
        ' .... Set the content to be signed
        oSignedData.Content = StrConv(oMsg.BodyPart.GetStream.ReadText, vbFromUnicode)
        
        ' .... Sign the content
        szSignature = oSignedData.Sign(oSigner, False, CAPICOM_ENCODE_BINARY)
        
        ' .... Get the string data as a byte array
        byteSignature = szSignature
        
        ' .... Attach the signature and let CDO base64 encode it
        Set oStream = oBodyPart.GetDecodedContentStream
        oStream.Type = ADODB.StreamTypeEnum.adTypeBinary
        oStream.Write (byteSignature)
        oStream.Flush
    End Select

    ' .... Signing Was sucessfull
    SignMessage = True
    Set oMsg = oSignedMsg
   
    MsgBox "MIME signing success!", vbInformation, App.Title
   
GoTo CleanUp

ErrorHandler:
    ' .... If the user cancels, don't display error message
    If Err.Number <> CAPICOM_E_CANCELLED Then
        MsgBox "Error: " & Hex(Err.Number) & ": " & Err.Description, vbExclamation, App.Title
    End If
    Err.Clear
    
    ' .... An error occurred
    SignMessage = False
    Set oMsg = Nothing

CleanUp:
    Set oSignedMsg = Nothing
    Set oBodyPart = Nothing
    Set cFields = Nothing
    Set oStream = Nothing
    Set oSignedData = Nothing
    Set oUtilities = Nothing
    Set oAttribute = Nothing
    Set oSignerCertificate = Nothing
    Set cSignerCertificates = Nothing
    Set oStore = Nothing
End Function

Private Function VerifyMessage(ByRef oInMsg As CDO.Message) As Boolean
'******************************************************************************
'
' Function:     VerifyMessage
'
' Parameters:   oMsg        -   A CDO object representing a properly formed S/MIME
'                               message. [in/out]
'
'
' Purpose:      Verify that a S/MIME message is signature valid
'
'******************************************************************************
    Dim oSignedData As New CAPICOM.SignedData
    Dim szSignature As String
    Dim iStart As Integer, iEnd As Integer, szTemp As String
    On Error GoTo ErrorHandler
    
    ' .... Get the pkcs7 signature
    szSignature = GetSignature(oInMsg)

    ' .... Verify The message
    Set oSignedData = New CAPICOM.SignedData
    
    ' .... Is this a detached or attached signature, deal with the differences
    oSignedData.Content = StrConv(GetContent(oInMsg), vbFromUnicode)
    Call oSignedData.Verify(szSignature, True, CAPICOM_VERIFY_SIGNATURE_ONLY)
    
    ' .... Update the global signer for use later
    Set oSigner = oSignedData.Signers.Item(1)
        
    VerifyMessage = True

GoTo CleanUp
ErrorHandler:
    MsgBox "Error: " & Hex(Err.Number) & ": " & Err.Description, vbExclamation, App.Title
    Err.Clear
    
    VerifyMessage = False

CleanUp:
    ' .... Clean up
    Set oSignedData = Nothing
End Function
