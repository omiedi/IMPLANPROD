VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form CONTENED 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contenedor Temporario "
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTCONTENEDOR 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   600
      Width           =   13740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   8
      Top             =   30
      Width           =   2175
   End
   Begin VB.TextBox TXTRUTA 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   0
      Width           =   11460
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   13800
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pintar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Pintar/Seleccionar Texto"
         Top             =   3240
         Width           =   650
      End
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
         Left            =   120
         Picture         =   "CONTENED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela Observaciones"
         Top             =   240
         Width           =   650
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
         Picture         =   "CONTENED.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba"
         Top             =   3960
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CONTENED.frx":0454
         Top             =   4680
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CONTENED.frx":2376
      Top             =   600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "CONTENED.frx":25AA
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   210
      TabIndex        =   4
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
      TabIndex        =   0
      Top             =   5250
      Width           =   795
   End
End
Attribute VB_Name = "CONTENED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
   TXTCONTENEDOR.SelStart = 0
   TXTCONTENEDOR.SelLength = Len(TXTCONTENEDOR.TEXT)
   TXTCONTENEDOR.SetFocus
   Clipboard.SetText Mid$(TXTCONTENEDOR.TEXT, TXTCONTENEDOR.SelStart + 1, TXTCONTENEDOR.SelLength)
   Archi$ = TXTRUTA.TEXT
   If TRIM$(Archi$) <> "" Then
     Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
   End If
End Sub

Private Sub Command14_Click()
   Label1 = "OK"
   Hide
End Sub

Private Sub Command2_Click()
   TXTCONTENEDOR.SelStart = 0
   TXTCONTENEDOR.SelLength = Len(TXTCONTENEDOR.TEXT)
   TXTCONTENEDOR.SetFocus

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
   Call APLICACIONSKINS(Me, Picture5)
   '
End Sub

Private Sub TXTCONTENEDOR_Change()
   Call PINTATEXT(CONTENED.TXTCONTENEDOR)

End Sub
