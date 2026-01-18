VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form BLOCNOTA 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Block de Notas"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7140
      Left            =   8160
      ScaleHeight     =   7080
      ScaleWidth      =   825
      TabIndex        =   1
      Top             =   0
      Width           =   885
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "BLOCNOTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Guardar"
         Top             =   4320
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "BLOCNOTA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   420
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   5280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "BLOCNOTA.frx":0454
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   1410
      End
   End
   Begin VB.TextBox Text1 
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "BLOCNOTA.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "BLOCNOTA.frx":25AA
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "BLOCNOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHX$

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()
'
    Call SaveFile(LUDAT$ + ARCHX$ + ".txt", Text1.TEXT)
'
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
    
    ARCHX$ = OPRECEP07.Text10.TEXT
    Text1.TEXT = LOADFILE(LUDAT$ + ARCHX$ + ".txt")
    '
End Sub
