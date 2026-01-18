VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ENTRAOBS 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Obs Pred."
      Height          =   260
      Left            =   4080
      TabIndex        =   16
      ToolTipText     =   "Observaciones Predefiniidas"
      Top             =   3660
      Width           =   960
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   10080
      ScaleHeight     =   8490
      ScaleWidth      =   1035
      TabIndex        =   9
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
         Height          =   650
         Left            =   130
         Picture         =   "ENTRAOBS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancela Observaciones"
         Top             =   100
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
         Height          =   810
         Left            =   130
         Picture         =   "ENTRAOBS.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aprueba, Observaciones"
         Top             =   3045
         Width           =   600
      End
      Begin VB.CommandButton Command13 
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
         Height          =   650
         Left            =   130
         Picture         =   "ENTRAOBS.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   600
      End
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   105
      TabIndex        =   7
      Top             =   3660
      Width           =   3870
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Selecc Todo"
      Height          =   260
      Left            =   8800
      TabIndex        =   6
      Top             =   3660
      Width           =   1080
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pegar"
      Height          =   260
      Left            =   7605
      TabIndex        =   5
      Top             =   3660
      Width           =   990
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      Height          =   260
      Left            =   6360
      TabIndex        =   4
      Top             =   3660
      Width           =   960
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cortar"
      Height          =   260
      Left            =   5175
      TabIndex        =   3
      Top             =   3660
      Width           =   960
   End
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   9780
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   105
      ScaleHeight     =   405
      ScaleWidth      =   9750
      TabIndex        =   0
      Top             =   210
      Width           =   9780
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   9570
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ENTRAOBS.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "ENTRAOBS.frx":0992
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   13
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
      TabIndex        =   8
      Top             =   5250
      Width           =   795
   End
End
Attribute VB_Name = "ENTRAOBS"
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

Private Sub Form_Load()
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture5)
   '
End Sub
