VERSION 5.00
Begin VB.Form SELDESTI09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Prospectos"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Selecciona Base Completa"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2730
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Clasificaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2150
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3990
      Begin VB.ListBox List5 
         Height          =   1635
         Left            =   130
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   360
         Width           =   3690
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "SELDESTI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label1 = "OK"
    Hide
End Sub

Private Sub Form_Load()
   List5.Clear
   For JJ& = 1 To ULTREG&("TAGRUCON")
      XX$ = REGLEIDO$("TAGRUCON", JJ&)
      RAGRE% = CVI(Left$(XX$, 2))
      TAGRE$ = FORMATNUM$(RAGRE%, "I3") + "  " + AJUSTI$(Mid$(XX$, 3, 48), 67) + Str$(RAGRE%)
      List5.AddItem TAGRE$
   Next JJ&
   Check1.Value = 0
   Label1 = ""
End Sub
