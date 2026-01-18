VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OBSERVAC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3840
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   6480
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   130
         Picture         =   "OBSERVAC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancela Observaciones"
         Top             =   50
         Width           =   600
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "OBSERVAC.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba, Observaciones"
         Top             =   1440
         Width           =   600
      End
   End
   Begin VB.TextBox Text28 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   120
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6300
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   105
      ScaleHeight     =   735
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   -30
      Width           =   6320
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6300
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OBSERVAC.frx":0454
      Top             =   600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OBSERVAC.frx":0688
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Caracteres Disponibles:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "512"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   2040
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   5250
      Width           =   795
   End
End
Attribute VB_Name = "OBSERVAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
   Label1 = "OK"
   Hide
End Sub

Private Sub Command22_Click()
   Unload Me
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
'
Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
'
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
'
Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
'

Private Sub Form_Activate()
   Call PINTATEXT(Text28)
End Sub

Private Sub Form_Load()
   '
   
   'Call APLICACIONSKINS(Me, Picture5)
   '
End Sub

Private Sub Text28_Change()
   
   Label2 = Text28.MaxLength - Len(Text28)
End Sub
