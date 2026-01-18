VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MODIGRADO1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Editar Matrícula"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   900
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   900
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.TextBox TXTTURNO 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   1670
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   975
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Combo8"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   340
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Turno:"
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.TextBox TXTPARA 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4080
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox ma 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   350
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Matriculado para:"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2115
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "Combo8"
         Top             =   240
         Width           =   1935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MODIGRADO1.frx":0000
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   8280
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TXTLETRA 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   255
         Width           =   510
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   340
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Letra:"
         Top             =   240
         Width           =   525
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   13
         Text            =   "Combo3"
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "MODIGRADO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CANCELA As Boolean


Private Sub Combo1_Change()
    TXTTURNO = Combo1.TEXT
End Sub

Private Sub Combo1_Click()
    TXTTURNO = Combo1.TEXT
End Sub

Private Sub Combo2_Change()
   TXTPARA = Combo2.TEXT
End Sub

Private Sub Combo2_Click()
    TXTPARA = Combo2.TEXT
End Sub

Private Sub Combo3_Change()
   TXTLETRA = Combo3.TEXT
End Sub

Private Sub Combo3_Click()
   TXTLETRA = Combo3.TEXT
End Sub

Private Sub Command1_Click()
   CANCELA = True
   Hide
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    If XVALO(Text3) < 1 Then
      MsgBox "FALTA INGRESAR NUMERO DE: " & TXTPARA.TEXT
      GoTo 99
    End If
    
    CANCELA = False
99  Command2.Enabled = True
    Hide
End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    '
    CANCELA = False
    '
    Combo2.AddItem "AÑO"
    Combo2.AddItem "GRADO"
    Combo2.AddItem "SALA DE"
    '
    Combo1.AddItem "MAÑANA"
    Combo1.AddItem "TARDE"
    Combo1.AddItem "NOCHE"
    '
    Combo3.AddItem "A"
    Combo3.AddItem "B"
    Combo3.AddItem "C"

End Sub
