VERSION 5.00
Begin VB.Form NOVMOD09 
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Inspección de Modelos"
   ClientHeight    =   5475
   ClientLeft      =   1890
   ClientTop       =   6375
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6180
      Left            =   7515
      ScaleHeight     =   6120
      ScaleWidth      =   1395
      TabIndex        =   21
      Top             =   -60
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         Picture         =   "NOVMOD09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Registra inspeccion de modelo"
         Top             =   3915
         Width           =   700
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         Picture         =   "NOVMOD09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Imprimir"
         Top             =   3255
         Width           =   700
      End
      Begin VB.CommandButton Command9 
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
         Height          =   660
         Left            =   120
         Picture         =   "NOVMOD09.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cierra y Sale de la Aplicación"
         Top             =   360
         Width           =   700
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -255
         Picture         =   "NOVMOD09.frx":0556
         Top             =   4920
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -435
         Picture         =   "NOVMOD09.frx":2478
         Top             =   7860
         Width           =   2010
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   330
      Left            =   5000
      TabIndex        =   11
      Top             =   315
      Width           =   2535
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
         Left            =   2205
         TabIndex        =   13
         Top             =   0
         Width           =   175
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   14
         Top             =   50
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Tipo de Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4725
      TabIndex        =   6
      Top             =   1080
      Width           =   2640
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Reparación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   8
         Top             =   570
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Inspección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   210
      TabIndex        =   0
      Top             =   2205
      Width           =   7170
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   7095
         TabIndex        =   2
         Top             =   0
         Width           =   7150
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle / Observaciones / Resultados "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   4
            Top             =   75
            Width           =   10200
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   420
         Width           =   7155
      End
      Begin VB.Line Line5 
         X1              =   7160
         X2              =   7155
         Y1              =   0
         Y2              =   3045
      End
      Begin VB.Line Line7 
         X1              =   10480
         X2              =   10480
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8400
         TabIndex        =   5
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   7170
         Y1              =   3060
         Y2              =   3060
      End
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
      Height          =   465
      Left            =   210
      TabIndex        =   20
      Top             =   1575
      Width           =   3375
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   1785
      TabIndex        =   19
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pieza:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   -210
      TabIndex        =   18
      Top             =   1290
      Width           =   1905
   End
   Begin VB.Label Label6 
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
      Height          =   465
      Left            =   210
      TabIndex        =   17
      Top             =   630
      Width           =   3375
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   1785
      TabIndex        =   16
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Ltota 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   -210
      TabIndex        =   15
      Top             =   350
      Width           =   1905
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   4725
      TabIndex        =   10
      Top             =   705
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0E0&
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   735
      Width           =   1275
   End
End
Attribute VB_Name = "NOVMOD09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Command1.Enabled = False
   'Call MENSERR(24, "No Disponible")
   Command1.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    FF% = FECHANUM(Text10)
    If FF% < 1 Or FF% > HOY(HOS$) Then
      Call MENSERR(24, "Fecha de Inspección no Válida.")
      GoTo 99
    End If
    '
    If TRIM$(Text1) = "" Then
      Call MENSERR(24, "No se Ha Ingresado\Detalle de Inspección / Reparación")
      GoTo 99
    End If
    '
    CODMODE$ = AJUSTI$(Label3, 16)
'    X$ = FILTERGETRECORD("BADAMODE", 1, CODMODE$)
    ALTMOD09.Label6 = Text10
'    FRECUENCIA% = Val(Mid$(X$, 97, 16))
    CONDI$ = "COD_MOD = '" & TRIM$(CODMODE$) & "'"
    FRECUENCIA% = XVALO(VALOBADA("BADAMODE", "FRECUEN", CONDI$))
    Select Case FRECUENCIA%
      Case 1: Call ALTMOD09.CALPROCON(7)
      Case 2: Call ALTMOD09.CALPROCON(30)
      Case 3: Call ALTMOD09.CALPROCON(90)
      Case 4: Call ALTMOD09.CALPROCON(180)
      Case 5: Call ALTMOD09.CALPROCON(365)
    End Select
    PROCON% = FECHANUM(ALTMOD09.Label9)
    ULTCON% = FECHANUM(ALTMOD09.Label6)
    'Call REPLA(X$, MKI$(PROCON%), 115, 2)
    'Call FILTERPUTRECORD("BADAMODE", 1, CODMODE$, X$)
    Call ACTUREGISTRO("BADAMODE", "FECHULT_CON", CONDI$, FETEXCOR1$(ULTCON%), REG&)
    Call ACTUREGISTRO("BADAMODE", "FECHPRO_CON", CONDI$, FETEXCOR1$(PROCON%), REG&)
    '
    TTXX$ = Space$(76)
    TIREGI$ = "Inspección": If Option2.Value = True Then TIREGI$ = "Reparación"
    Call REPLA(TTXX$, "Registro de " + TIREGI$ + ":", 11, 56)
    Call REPLA(TTXX$, HOS$, 1, 8)
    Call REPLA(TTXX$, USERNAME$, 69, 24)
    ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
    Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
    TTXX$ = Space$(76)
    ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
    Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
    '
    Call FRATEXTO(Text1, 56)
    For KKI% = 1 To CARETEX%
      TTXX$ = Space$(76)
      Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
      ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
      Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
    Next KKI%
    '
    TTXX$ = Space$(76)
    ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
    Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
    TTXX$ = Space$(76)
    FEX = PROCON%
    Call REPLA(TTXX$, "Próxima Inspección: " + FECHATEX$(FEX), 11, 56)
    ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
    Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
    '
    ' Call REGAPP("ANOTAMOD", CODMODE$ + String$(88, "="))
    Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, String$(88, "="))
    ' Call CIERRARCH("ANOTAMOD")
    Text1 = ""
    '
    Command3.Enabled = True
    Hide
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Text10_DblClick()
   Call Command5_Click
End Sub

