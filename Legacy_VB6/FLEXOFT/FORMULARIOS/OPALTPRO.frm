VERSION 5.00
Begin VB.Form OPALTPRO 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Altas y Modificaciones de Proveedores"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7245
      ScaleHeight     =   75
      ScaleWidth      =   690
      TabIndex        =   33
      Top             =   3465
      Width           =   750
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0C0&
      Caption         =   "I.B.Pcia.Bs.As"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   5240
      TabIndex        =   30
      Top             =   3390
      Width           =   1800
      Begin VB.CommandButton Command3 
         Height          =   540
         Left            =   210
         Picture         =   "OPALTPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Configuracion Funcionamiento Retenciones IB Pcia de Bs.As"
         Top             =   280
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Height          =   540
         Left            =   945
         Picture         =   "OPALTPRO.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Listado Proveedores para I.Brutos"
         Top             =   280
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Controles Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Left            =   210
      TabIndex        =   4
      Top             =   3360
      Width           =   6840
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Gan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   4200
         TabIndex        =   35
         Top             =   35
         Width           =   855
         Begin VB.CommandButton Command5 
            Height          =   540
            Left            =   105
            Picture         =   "OPALTPRO.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Configuracion Funcionamiento Retenciones de Ganancias"
            Top             =   280
            Width           =   645
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Revisar Todos los C.U.I.T.s al Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   500
         TabIndex        =   5
         Top             =   420
         Width           =   3480
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Valores Preteterminados por Omisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   6840
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   0
         Left            =   2310
         TabIndex        =   29
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1785
         TabIndex        =   27
         Text            =   "  "
         Top             =   1155
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   2625
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1155
         Width           =   4035
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   2625
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1785
         Width           =   4035
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   2625
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1470
         Width           =   4035
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   1785
         TabIndex        =   18
         Text            =   " "
         Top             =   1785
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   1785
         TabIndex        =   17
         Text            =   " "
         Top             =   840
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   2625
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   840
         Width           =   4035
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   1785
         TabIndex        =   15
         Top             =   525
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   2625
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   525
         Width           =   4035
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   1785
         TabIndex        =   13
         Text            =   " "
         Top             =   1470
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   4240
         TabIndex        =   12
         Text            =   " "
         Top             =   2310
         Width           =   2250
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
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
         Index           =   5
         Left            =   1785
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   2625
         Width           =   4875
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   1
         Left            =   2310
         TabIndex        =   10
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   2
         Left            =   2310
         TabIndex        =   9
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   3
         Left            =   2310
         TabIndex        =   8
         Top             =   1470
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   4
         Left            =   2310
         TabIndex        =   7
         Top             =   1785
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   280
         Index           =   5
         Left            =   6470
         TabIndex        =   6
         Top             =   2310
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
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
         Left            =   525
         TabIndex        =   28
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Madre:"
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
         Left            =   2205
         TabIndex        =   25
         Top             =   2415
         Width           =   1905
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cond.Pago:"
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
         Left            =   0
         TabIndex        =   24
         Top             =   1890
         Width           =   1710
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría I-B"
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
         Left            =   315
         TabIndex        =   23
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.IVA"
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
         Left            =   315
         TabIndex        =   22
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta:"
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
         Left            =   0
         TabIndex        =   21
         Top             =   1575
         Width           =   1710
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7245
      Picture         =   "OPALTPRO.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1785
      Width           =   660
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7245
      Picture         =   "OPALTPRO.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   660
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
      Height          =   700
      Left            =   7245
      Picture         =   "OPALTPRO.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6615
      Picture         =   "OPALTPRO.frx":164C
      Top             =   3800
      Width           =   2010
   End
End
Attribute VB_Name = "OPALTPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ESCALRET() 'ESCALAS DE RETENCION DE GANANCIAS
'\\\OT-06-0559-RG-18/12/06///
  '
  Call SELECHO("TACRETG")
  If TRIM$(VALACT1$("TACRETG")) <> "" Then
    '3.- SI NO TIENE MINIMO O TASA => AFUERA
    NUCA% = XVALO(VALACT1$("TACRETG")) 'CATEGORIA ELEGIDA
    X$ = FILTERGETRECORD$("TACRETG", 1, Chr(NUCA%))
    DESCRICAT$ = Mid$(X$, 2, 31)
    MINEX1 = CVD(Mid$(X$, 33, 8))
    TASA1 = CVS(Mid$(X$, 41, 4))
    If MINEX1 <= 0.005 And TASA1 <= 0.005 Then
      Call COMUNI("No se puede Cargar Escala a Categoria Exento")
      Exit Sub
    End If
    '6.- CARGA LA TABLA C/REGISTROS Q' MACHEEN CON LA CATEGORIA SELECIONADA
    Call FILTERFILE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    '7.-GRABACION
    VTB% = VENTABU%("ESCALGAN/NC/TITUPAN=Escala de Retencion - " & DESCRICAT$ & "")
    If VTB% >= 0 Then
      For J& = 1 To ULTREG&("!ESCALGAN")
        YY$ = REGLEIDO$("!ESCALGAN", J&)
        Call REPLA(YY$, MKI$(NUCA%), 1, 2)
        Call GRAREG("!ESCALGAN", YY$, J&)
      Next J&
      Call FILTERUPDATE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    End If
    '
  End If
  '
'\\\OT-06-0559-RG-FIN///
End Sub

Sub CARTARET()
'\\\OT-06-0559-RG-18/12/06///
    Call TRALOCAL("TACRETG", "")
     VTB% = VENTABU%("TACRETG")
     If VTB% >= 0 Then
       For U& = 1 To ULTREG&("!TACRETG")
         XX$ = REGLEIDO$("!TACRETG", U&)
         ' Graba el numero de orden automaticamente
         Call REPLA(XX$, Chr$(U& - 1), 1, 1)
         Call GRAREG("TACRETG", XX$, U&)
       Next U&
       Call CIERRARCH("TACRETG")
     End If
'\\\OT-06-0559-RG-FIN///
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    For Index = 0 To 5
       TV11% = TIPFIELD(ARCHE$(Index), 1)
       LF11% = LENFIELD(ARCHE$(Index), 1)
       ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
       VAPRE$ = "": ECO$ = ""
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
             VAPRE$ = TRIM$(COTEXT(Index).TEXT)
          End If
       End If
       COTEXT(Index).TEXT = VAPRE$
       DETEXT(Index).TEXT = ECO$
       Call GRACONTROL("AMAPRO", ARCHE$(Index), VAPRE$)
    Next Index
    '
    MONUME$ = "NO": VERICUIT% = (-1)
    If Check1.Value = 1 Then
       MONUME$ = "SI": VERICUIT% = 1
    End If
    Call GRACONTROL("AMAPRO", "VERICUIT", MONUME$)
    '
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parámetros Ingresos Brutos") = 2 Then
     If ULTREG&("TACATIBP") < 5 Then
       Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
       Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
       Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
       Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
       Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
       Call CIERRARCH("TACATIBP")
     End If
     '
     Call TRALOCAL("TACATIBP", "")
     VTB% = VENTABU%("TACATIBP")
     If VTB% >= 0 Then
       For U& = 1 To ULTREG&("!TACATIBP")
         XX$ = REGLEIDO$("!TACATIBP", U&)
         Call REPLA(XX$, Chr$(U& - 1), 1, 1)
         Call GRAREG("TACATIBP", XX$, U&)
       Next U&
       Call CIERRARCH("TACATIBP")
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call FILESORT("DARIBPRO", "DARIBPRJ", 1, 1)
   Screen.MousePointer = 11
   FIN& = ULTREG&("DARIBPRJ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DARIBPRJ", U&)
     NCLIX% = CVI(Left$(XX$, 2))
     CATIBX% = Val(CATIBRU$((-1) * NCLIX%))
     NIBR$ = NIBRCLI$((-1) * NCLIX%)
     Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
     Call REPLA(XX$, NIBR$, 67, 16)
     If CATIBX% > 0 Then
       If TRIM$(NIBR$) = "" Then
         Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
       End If
     End If
     Call GRAREG("DARIBPRJ", XX$, U&)
   Next U&
   Call CIERRARCH("DARIBPRJ")
   Screen.MousePointer = 1
   '
   Call FINAL("*LIDARIBP")
   '
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  '\\\OT-06-0559-RG-18/12/06///
  Command5.Enabled = False
    RTA% = ALTERNA%("Tabla de Categorias\Escalas de Retención")
    If RTA% = 1 Then
        Call CARTARET
      ElseIf RTA% = 2 Then
        Call ESCALRET
    End If
  Command5.Enabled = True
  '\\\OT-06-0559-RG-FIN///
End Sub


Private Sub Form_Load()
    '
    For Index = 0 To 5
       VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
       COTEXT(Index) = "": DETEXT(Index) = ""
       If VAPRE$ <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
              COTEXT(Index).TEXT = VAPRE$
              DETEXT(Index).TEXT = ECO$
            Else
              COTEXT(Index).TEXT = ""
              DETEXT(Index).TEXT = ""
          End If
       End If
       DETEXT(Index).Refresh
       DETEXT(Index).TabStop = False
    Next Index
    '
    VERICUIT% = 1: Check1.Value = 1
    If CONTROL$("AMAPRO", "VERICUIT") = "NO" Then
       VERICUIT% = (-1)
       Check1.Value = 0
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Then
      Frame3.Visible = True
    End If
    '
End Sub
'

Private Sub COTEXT_DblClick(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub

