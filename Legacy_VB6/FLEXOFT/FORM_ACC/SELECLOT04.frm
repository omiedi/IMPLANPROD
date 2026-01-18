VERSION 5.00
Begin VB.Form SELECLOT04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección por Lote"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   2100
      Left            =   5985
      ScaleHeight     =   2040
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   2200
         Left            =   0
         Picture         =   "SELECLOT04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6000
      Left            =   5985
      ScaleHeight     =   5940
      ScaleWidth      =   2160
      TabIndex        =   13
      Top             =   0
      Width           =   2220
      Begin VB.CommandButton Command29 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   105
         Picture         =   "SELECLOT04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aprueba la Selección"
         Top             =   1050
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "SELECLOT04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2835
      TabIndex        =   3
      Top             =   1575
      Width           =   175
   End
   Begin VB.Label Label8 
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
      Left            =   5250
      TabIndex        =   12
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   11
      Top             =   285
      Width           =   1020
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
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
      Left            =   4515
      TabIndex        =   10
      Top             =   1575
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote o Partida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   9
      Top             =   1470
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   4515
      TabIndex        =   8
      Top             =   1050
      Width           =   1230
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -840
      TabIndex        =   7
      Top             =   400
      Width           =   1965
   End
   Begin VB.Label Label2 
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
      Left            =   420
      TabIndex        =   6
      Top             =   630
      Width           =   5325
   End
   Begin VB.Label Label1 
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
      Left            =   2310
      TabIndex        =   5
      Top             =   210
      Width           =   1860
   End
   Begin VB.Label Label28 
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
      Left            =   1155
      TabIndex        =   4
      Top             =   1575
      Width           =   1860
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   2
      Top             =   1125
      Width           =   1965
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   315
      TabIndex        =   1
      Top             =   255
      Width           =   1965
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   1620
      Width           =   1965
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "SELECLOT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call Command29_Click
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    If CODINT%(Label1) < 1 Then
       Call MENSERR(24, "Código Inexistente o no Válido")
       GoTo 99
    End If
    '
    If TRIM$(Label28) = "" Then
      Call MENSERR(24, "No se Ha Elegido el Lote")
      GoTo 99
    End If
    '
    If XVALO(Label5) > XVALO(Label7) + 0.05 Then
       Call MENSERR(24, "Saldo Insuficiente en el Lote Elegido.")
       GoTo 99
    End If
    '
    Command29.Enabled = True
    Hide
    Exit Sub
    '
99  Command29.Enabled = True
End Sub

Private Sub Command5_Click()
    CIXI% = CODINT%(Label1)
    DEPX% = XVALO(Label4)
    Call ARMALILOT(CIXI%)
    Screen.MousePointer = 1
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    Label28 = VDEVU$
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
    End If

End Sub

Private Sub Label1_Change()
    CICOMPO% = CODINT%(Label1)
    Label2 = DESCRIT0$(CICOMPO%)
    Label8 = UNIMED(CICOMPO%)
End Sub

Private Sub Label28_Change()
    Label7 = ""
    CIXI% = CODINT%(Label1)
    IDLO$ = TRIM$(Label28)
    DEPX% = XVALO(Label4)
    Label7 = ""
    If TRIM$(IDLO$) <> "" Then
       Label7 = FORMATNUM$(SALOTE(CIXI%, IDLO$), "F10.1")
    End If
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub
