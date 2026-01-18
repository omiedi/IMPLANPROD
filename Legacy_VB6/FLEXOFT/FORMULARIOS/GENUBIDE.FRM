VERSION 5.00
Begin VB.Form GENUBIDE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Ubicaciones"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   540
      Left            =   3885
      Picture         =   "GENUBIDE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   1470
      Width           =   645
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   945
      Width           =   750
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   945
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   945
      Width           =   750
   End
   Begin VB.CommandButton Command9 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1740
      TabIndex        =   1
      Top             =   195
      Width           =   175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1215
      TabIndex        =   0
      Top             =   195
      Width           =   540
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   11
      Top             =   1515
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1215
      TabIndex        =   10
      Top             =   1470
      Width           =   2565
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3150
      TabIndex        =   8
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Columna:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1890
      TabIndex        =   6
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pasillo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2055
      TabIndex        =   3
      Top             =   195
      Width           =   2430
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   2
      Top             =   245
      Width           =   915
   End
End
Attribute VB_Name = "GENUBIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
   If TRIM$(Label4) = "" Then
      Call MENSERR(24, "Ubicacion no Válida")
      Exit Sub
   End If
   '
   Hide
End Sub

Private Sub text6_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command9_Click
   End If
End Sub

Private Sub Text6_Change()
   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

Private Sub Combo1_CLICK()
   Call CALCUBI
End Sub

Private Sub Combo2_CLICK()
   Call CALCUBI
End Sub

Private Sub Combo3_CLICK()
   Call CALCUBI
End Sub

Private Sub Command9_Click()
   Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Form_Load()
   Combo1.Clear
   For U& = 65 To 90
      Combo1.AddItem Chr$(U&)
   Next U&
   For U& = 11 To 19
      Combo2.AddItem TRIM$(Str$(U&))
   Next U&
   For U& = 1 To 6
      Combo3.AddItem TRIM$(Str$(U&))
   Next U&

End Sub
'
Sub CALCUBI()
   UBICA$ = TRIM$(Text6.TEXT)
   If UBICA$ <> "" Then UBICA$ = UBICA$ + "-"
   UBICA$ = UBICA$ + Combo1.TEXT + "-" + Combo2.TEXT + "-" + Combo3.TEXT
   '
   If Combo1.TEXT = "" Then UBICA$ = ""
   If Combo2.TEXT = "" Then UBICA$ = ""
   If Combo3.TEXT = "" Then UBICA$ = ""
   '
   Label4 = UBICA$
End Sub
