VERSION 5.00
Begin VB.Form RESCUEDIAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Mayor por Cuenta"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
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
   ScaleHeight     =   4125
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
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
         Height          =   855
         Left            =   8640
         Picture         =   "RESCUEDIAR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Height          =   700
         Left            =   8610
         Picture         =   "RESCUEDIAR.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   3255
         Width           =   855
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
         Height          =   855
         Left            =   8610
         Picture         =   "RESCUEDIAR.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1365
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dispositivo de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   24
         Top             =   315
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Pantalla"
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
            Index           =   1
            Left            =   735
            TabIndex        =   26
            Top             =   420
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Impresora"
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
            Index           =   2
            Left            =   2205
            TabIndex        =   25
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Rango de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   4410
         TabIndex        =   15
         Top             =   315
         Width           =   4005
         Begin VB.TextBox Text7 
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
            Left            =   315
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1785
            Width           =   3330
         End
         Begin VB.CommandButton Command8 
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
            Height          =   360
            Left            =   3465
            TabIndex        =   20
            Top             =   1365
            Width           =   175
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1995
            TabIndex        =   19
            Text            =   " "
            Top             =   1365
            Width           =   1485
         End
         Begin VB.TextBox Text9 
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
            Left            =   315
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   " "
            Top             =   840
            Width           =   3330
         End
         Begin VB.CommandButton Command9 
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
            Height          =   360
            Left            =   3465
            TabIndex        =   17
            Top             =   420
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
            Height          =   360
            Left            =   1995
            TabIndex        =   16
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta :"
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
            Left            =   -420
            TabIndex        =   23
            Top             =   1470
            Width           =   2325
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   -420
            TabIndex        =   22
            Top             =   525
            Width           =   2325
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Avance de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   210
         TabIndex        =   12
         Top             =   3150
         Width           =   4005
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   315
            ScaleHeight     =   75
            ScaleWidth      =   3315
            TabIndex        =   13
            Top             =   420
            Width           =   3375
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   12000
            End
         End
      End
      Begin VB.ListBox List1 
         Height          =   540
         Left            =   8610
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2415
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   210
         TabIndex        =   4
         Top             =   1470
         Width           =   4005
         Begin VB.CommandButton Command6 
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
            Height          =   360
            Left            =   2730
            TabIndex        =   8
            Top             =   420
            Width           =   175
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   7
            Top             =   420
            Width           =   1065
         End
         Begin VB.CommandButton Command10 
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
            Height          =   360
            Left            =   2730
            TabIndex        =   6
            Top             =   945
            Width           =   175
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   5
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   -735
            TabIndex        =   10
            Top             =   525
            Width           =   2325
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   735
            TabIndex        =   9
            Top             =   1050
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1120
         Left            =   4410
         TabIndex        =   1
         Top             =   2835
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Todas las Cuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   2415
            TabIndex        =   3
            Top             =   525
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Cuentas con Movimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   420
            TabIndex        =   2
            Top             =   525
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "RESCUEDIAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload RESCUEDIAR
End Sub
'

Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call INDEXCAJA(MAXDIA&)
   Call INDEXDIARIO(MAXDIARIO%)
   If Option1(1) = True Then
        Call LIMAYORPAN
      Else
        Call LIMAYORPRIN
   End If
99 Command2.Enabled = True
End Sub
'

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECON")
    ULTCUEX$ = VALACT1$("ECUECON")
    If ULTCUEX$ <> "" Then
        Text8.TEXT = ULTCUEX$
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECON")
    PRICUEX$ = VALACT1$("ECUECON")
    If PRICUEX$ <> "" Then
        Text10.TEXT = PRICUEX$
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Dim V0 As String * 32
    LU$ = LUDAT$
    N9% = FILEOPEN%(LU$ + "VARIAB." + EXTE$, 0, 32)
    Get #N9%, 12 + OPCI%, V0$
    MODO% = Asc(Left$(V0$, 1))
    DES% = CVI(Mid$(V0$, 2, 2))
    HAS% = CVI(Mid$(V0$, 4, 2))
    NIPRICUE% = CVI(Mid$(V0$, 6, 2))
    NIULTCUE% = CVI(Mid$(V0$, 8, 2))
    SELE% = Asc(Mid$(V0$, 10, 1))
    '
    If MODO% < 1 Or MODO% > 2 Then MODO% = 1
    If SELE% < 1 Or SELE% > 2 Then SELE% = 1
    VMODO% = MODO%: VSELE% = SELE%
    '
    Get #N9%, 2, V0$
    V8$ = V0$
    EDES% = CVI(Mid$(V8$, 29, 2))
    EHAS% = CVI(Mid$(V8$, 31, 2))
    If DES% = 0 Then DES% = EDES%
    If HAS% = 0 Then HAS% = EHAS%
    FEX = DES%:  DESX$ = FECHATEX$(FEX)
    FEX = HAS%:   HASX$ = FECHATEX$(FEX)
    EJU$ = TRIM$(Mid$(V8$, 1, 28))
    '
    If NIPRICUE% = 0 Then
       XX$ = REGLEIDO$("INVECUE", 1)
       NIPRICUE% = CVI(Mid$(XX$, 13, 2))
    End If
    PRICUEX$ = CODCUE$(NIPRICUE%)
    '
    If NIULTCUE% = 0 Then
       XX$ = REGLEIDO$("INVECUE", ULTREG&("INVECUE"))
       NIULTCUE% = CVI(Mid$(XX$, 13, 2))
    End If
    ULTCUEX$ = CODCUE$(NIULTCUE%)
    '
    Text6.TEXT = DESX$
    Text5.TEXT = HASX$
    Text10.TEXT = PRICUEX$
    Text8.TEXT = ULTCUEX$
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub


Private Sub Text10_Change()
    Text9.TEXT = ECOARCH("ECUECON", Text10.TEXT)
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5.TEXT)
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
End Sub

Private Sub Text8_Change()
    Text7.TEXT = ECOARCH("ECUECON", Text8.TEXT)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECUECON")
    PRICUEX$ = VALACT1$("ECUECON")
    If PRICUEX$ <> "" Then
        Text10.TEXT = PRICUEX$
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("ECUECON")
    ULTCUEX$ = VALACT1$("ECUECON")
    If ULTCUEX$ <> "" Then
        Text8.TEXT = ULTCUEX$
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub LIMAYORPAN()
   '
   DES% = FECHANUM(Text6.TEXT)
   If DES% < COMEJ% Then
      DES% = COMEJ%
      FEX = DES%
      Text6.TEXT = FECHATEX$(FEX)
      Text6.Refresh
      Exit Sub
   End If
   '
   HAS% = FECHANUM(Text5.TEXT)
   FULTIAS% = FULTAS%
   If FULTASDIA% > FULTIAS% Then FULTIAS% = FULTASDIA%
   If HAS% > FULTIAS% Then
      HAS% = FULTIAS%
      FEX = HAS%
      Text5.TEXT = FECHATEX$(FEX)
      Text5.Refresh
      Exit Sub
   End If
   '
   If DES% > HAS% Then
      Call MENSERR(24, "Rango de Fechas Imposible")
      Exit Sub
   End If
   '
10 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then Exit Sub
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 10
   End If
   '
   Frame1.Refresh
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0: TRANSPORTE# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      Call LEEPASDIA(REGDIA%, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASDIA(REGDIA%, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 19
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
19 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   ASICAPEN# = 0
   For KKL& = 1 To Len(PMC$) Step 4
      UU& = KKL& + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KKL&, 4))
      Call LEESUCAJA(REGDIAL&)
      If Abs(CVS(D19$)) < 0.1 Then
         ASICAPEN# = ASICAPEN# + CVD(D16$)
      End If
29 Next KKL&
   '
   If Abs(ASICAPEN#) > 0.005 Then
      FECHASIEN% = HOY(HOS$)
      NROASIEN = 0
      NROASDEF% = 0
      MARCADIA$ = " "
      DETAPAS$ = "Asientos de Caja Pendientes"
      IMPOPAS# = ASICAPEN#
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 39
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 90
      End If
   End If
   '
39 If Abs(TRANSPORTE#) >= 0.005 Then
      FEX = DES%
      FECHASIEN% = DIANTE(FEX)
      NROASIEN = 0
      NROASDEF% = 0
      MARCADIA$ = " "
      DETAPAS$ = "SALDO ANTERIOR"
      IMPOPAS# = TRANSPORTE#
      GoSub 90
   End If
   '
   Call SETULTREG("RESUCON", JJ&)
   Call CIERRARCH("RESUCON")
   If JJ& < 1 Then
        Call MENSERR(24, "Cuenta " + CODENOCUE$ + "\sin Movimientos en el Ejercicio.")
      Else
         Call CONECRUN("*MAYOCON/" + CODENOCUE$, "Libro Mayor por Cuenta")
   End If
   GoTo 10
   '
90 XX$ = REGBLAN$("RESUCON")
   Call REPLA(XX$, CODENOCUE$, 1, 64)
   Call REPLA(XX$, MKI$(FECHASIEN%), 65, 2)
   Call REPLA(XX$, MKS$(NROASIEN), 67, 4)
   Call REPLA(XX$, MKI$(NROASDEF%), 71, 2)
   Call REPLA(XX$, MARCADIA$, 73, 2)
   Call REPLA(XX$, DETAPAS$, 75, 30)
   If IMPOPAS# >= 0 Then
         Call REPLA(XX$, MKD$(IMPOPAS#), 105, 8)
      Else
         Call REPLA(XX$, MKD$(Abs(IMPOPAS#)), 113, 8)
   End If
   JJ& = JJ& + 1
   Call GRAREG("RESUCON", XX$, JJ&)
   '
Return
'
End Sub

Sub LIMAYORPRIN()
   '
   Dim SUDEBIN#(16384)
   Dim SUHABIN#(16384)
   Dim SUDEBE#(16384)
   Dim SUHABE#(16384)
   Dim ASICAPE#(16384)
   Dim PUNTE$(16384)
   '
   DES% = FECHANUM(Text6.TEXT)
   If DES% < COMEJ% Then
      DES% = COMEJ%
      FEX = DES%
      Text6.TEXT = FECHATEX$(FEX)
      Text6.Refresh
   End If
   '
   HAS% = FECHANUM(Text5.TEXT)
   If HAS% > FULTASDIA% Then
      HAS% = FULTASDIA%
      FEX = HAS%
      Text5.TEXT = FECHATEX$(FEX)
      Text5.Refresh
   End If
   '
   If DES% > HAS% Then
      Call MENSERR(24, "Rango de Fechas Imposible")
      Exit Sub
   End If
   '
   DESCUE$ = Text10.TEXT
   If TRIM$(DESCUE$) = "" Then
      Call MENSERR(24, "Falta Rango de Cuentas")
   End If
   '
   HASCUE$ = Text8.TEXT
   If TRIM$(HASCUE$) = "" Then
      Call MENSERR(24, "Falta Rango de Cuentas")
   End If
   '
   If DESCUE$ > HASCUE$ Then
      Call MENSERR(24, "Rango de Cuentas Imposible")
      Exit Sub
   End If
   '
   INCLUCAJ% = Val(CONTROL$("CONTAGEN", "INCLUCAJ"))
   If INCLUCAJ% <> 1 Then INCLUCAJ% = 0
   RENUMERA% = Val(CONTROL$("CONTAGEN", "RENUMERA"))
   If RENUMERA% <> 1 Then RENUMERA% = 0
   '
   Screen.MousePointer = 11
   TODAS% = 0
   If Option1(0).Value = True Then TODAS% = 1
   '
   If INCLUCAJ% < 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEESUCAJA(KKL&)
       '
       CUE% = CVI(D11$)
       FEAS% = CVI(D15$)
       IPAS# = CVD(D16$)
       If CUE% > 0 Then
         If CUE% <= 16384 Then
           If FEAS% <= HAS% Then
             If Abs(CVS(D19$)) < 0.1 Then
               ASICAPE#(CUE%) = ASICAPE#(CUE%) + CVD(D16$)
             End If
           End If
         End If
       End If
     Next KKL&
   End If
   '
   Screen.MousePointer = 11
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEDIARIO(KK%)
      '
      If INCLUCAJ% > 0 Then
        If Asc(E18$) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA
      End If
      '
      CUE% = CVI(E11$)
      FEAS% = CVI(E15$)
      IPAS# = CVD(E16$)
      If CUE% > 0 Then
         If CUE% <= 16384 Then
            If FEAS% < DES% Then
                  If IPAS# > 0 Then
                      SUDEBIN#(CUE%) = SUDEBIN#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                     Else
                      SUHABIN#(CUE%) = SUHABIN#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                  End If
                Else
                  If FEAS% <= HAS% Then
                     If IPAS# > 0 Then
                         SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                        Else
                         SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                     End If
                     KKL& = KK%
                     PUNTE$(CUE%) = PUNTE$(CUE%) + MKS$(KKL&)
                  End If
            End If
         End If
      End If
      '
29 Next KK%
   '
   If INCLUCAJ% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEESUCAJA(KKL&)
       '
       CUE% = CVI(D11$)
       FEAS% = CVI(D15$)
       IPAS# = CVD(D16$)
       If CUE% > 0 Then
         If CUE% <= 16384 Then
            If FEAS% < DES% Then
                  If IPAS# > 0 Then
                      SUDEBIN#(CUE%) = SUDEBIN#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                     Else
                      SUHABIN#(CUE%) = SUHABIN#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                  End If
                Else
                  If FEAS% <= HAS% Then
                     If IPAS# > 0 Then
                         SUDEBE#(CUE%) = SUDEBE#(CUE%) + (Int(100 * IPAS# + 0.5)) / 100
                        Else
                         SUHABE#(CUE%) = SUHABE#(CUE%) - (Int(100 * IPAS# + 0.5)) / 100
                     End If
                     PUNTE$(CUE%) = PUNTE$(CUE%) + MKS$((-1) * KKL&)
                  End If
            End If
         End If
       End If
       '
     Next KKL&
   End If
   '
   If NX3% < 1 Then
      NX3% = FILEOPEN%(LUDAT$ + "INVECUE" + "." + EXTE$, 0, 16)
   End If
   '
   Screen.MousePointer = 11
   JJ& = 0
   Dim IVX As String * 16
   FIN& = LOF(NX3%) / 16
   For KK% = 1 To LOF(NX3%) / 16
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Get #NX3%, KK%, IVX$
      COCUE$ = TRIM$(Left$(IVX$, 12))
      '
      If COCUE$ >= DESCUE$ Then
        If COCUE$ <= HASCUE$ Then
          CUE% = CVI(Mid$(IVX$, 13, 2))
          If CUE% > 0 Then
            If CUE% <= 16384 Then
              If TODAS% = 1 Or Abs(SUDEBIN#(CUE%)) + Abs(SUHABIN#(CUE%)) + Abs(SUDEBE#(CUE%)) + Abs(SUHABE#(CUE%)) + Abs(ASICAPE#(CUE%)) >= 0.005 Then
                '
                List1.Clear
                SUTRA# = 0
                TTTT = Len(PUNTE$(CUE%))
                For KL% = 1 To Len(PUNTE$(CUE%)) Step 4
                  KML& = CVS(Mid$(PUNTE$(CUE%), KL%, 4))
                  If KML& > 0 Then
                     KM% = KML&
                     Call LEEDIARIO(Abs(KM%))
                     VALOSOR% = CVI(E15$)
                     If RENUMERA% = 1 Then VALOSOR% = CVI(E13$)
                     ZZ$ = "A" + CSTRING$(MKI$(VALOSOR%), 1, 6, 0, 2) + Space$(9)
                     Call REPLA(ZZ$, "A", 8, 1)
                     Call REPLA(ZZ$, Str$(KM%), 9, 8)
                     List1.AddItem ZZ$
                   ElseIf KML& < 0 Then
                     Call LEESUCAJA(Abs(KML&))
                     VALOSOR% = CVI(D15$)
                     If RENUMERA% = 1 Then VALOSOR% = CVI(D13$)
                     ZZ$ = "A" + CSTRING$(MKI$(VALOSOR%), 1, 6, 0, 2) + Space$(9)
                     Call REPLA(ZZ$, "A", 8, 1)
                     Call REPLA(ZZ$, Str$(KML&), 9, 8)
                     List1.AddItem ZZ$
                 End If
                Next KL%
                '
                If JJ& > 0 Then
                  XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                  XX$ = String$(128, "="): GoSub 95
                  XX$ = REGBLAN$(LIMAYOL$): GoSub 95
                End If
                '
                FECHASIEN% = 0: NROASIEN = 0: NROASDEF = 0: MARCADIA$ = "": REFE$ = "": IDEBE# = 0: IHABE# = 0
                REFE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%)): GoSub 90
                REFE$ = String$(Len(REFE$), "*"): GoSub 90
                XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                '
                FEX = DES
                FECHASIEN% = DIANTE(FEX)
                NROASIEN = 0
                NROASDEF = 0
                MARCADIA$ = " "
                REFE$ = "Transporte"
                IDEBE# = SUDEBIN#(CUE%)
                IHABE# = SUHABIN#(CUE%)
                GoSub 90
                '
                XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                '
                For KL% = 1 To List1.ListCount
                  KML& = Val(Mid$(List1.List(KL% - 1), 9))
                  If KML& > 0 Then
                     KM% = KML&
                     Call LEEPASDIA(KM%, 2)
                     FECHASIEN% = CVI(E15$)
                     NROASIEN = CVS(E12$)
                     NROASDEF% = CVI(E13$)
                     MARCADIA$ = "  "
                     REFE$ = TRIM$(DETAPAS$)
                     If REFE$ <> "" Then REFE$ = REFE$ + " - "
                     REFE$ = REFE$ + TRIM$(REFASIEN$)
                     IPAS# = CVD(E16$)
                     If IPAS# > 0 Then
                         IDEBE# = IPAS#: IHABE# = 0
                       Else
                         IDEBE# = 0: IHABE# = Abs(IPAS#)
                     End If
                     If Abs(IDEBE#) + Abs(IHABE#) > 0.005 Then GoSub 90
                   ElseIf KML& < 0 Then
                     Call LEEPASCAJ((Abs(KML&)), 2)
                     FECHASIEN% = CVI(D15$)
                     NROASIEN = CVS(D12$)
                     NROASDEF% = CVI(D13$)
                     MARCADIA$ = "  "
                     REFE$ = TRIM$(DETAPAS$)
                     If REFE$ <> "" Then REFE$ = REFE$ + " - "
                     REFE$ = REFE$ + TRIM$(REFASIEN$)
                     IPAS# = CVD(D16$)
                     If IPAS# > 0 Then
                         IDEBE# = IPAS#: IHABE# = 0
                       Else
                         IDEBE# = 0: IHABE# = Abs(IPAS#)
                     End If
                     If Abs(IDEBE#) + Abs(IHABE#) > 0.005 Then GoSub 90
                  End If
                Next KL%
                '
                If Abs(ASICAPE#(CUE%)) >= 0.005 Then
                   XX$ = REGBLAN$("LIMAYOL"): GoSub 95
                   FECHASIEN% = HAS%
                   NROASIEN = 0
                   NROASDEF% = 0
                   MARCADIA$ = "* "
                   REFE$ = "Asientos de Caja Pendientes"
                   IPAS# = ASICAPE#(CUE%)
                   If IPAS# > 0 Then
                        IDEBE# = IPAS#: IHABE# = 0
                       Else
                        IDEBE# = 0: IHABE# = Abs(IPAS#)
                    End If
                    GoSub 90
                End If
                '
              End If
            End If
          End If
        End If
      End If
      '
   Next KK%
   '
   Close #NX3%: NX3% = 0
   Call SETULTREG("LIMAYOL", JJ&)
   Call CIERRARCH("LIMAYOL")
   '
   Call HEADERS("LIMAYOL", "")
   Call HEADERS("LIMAYOL", "Corresponde a: " + EJACT$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*LIMAYOL", "Libro Mayor por Cuenta")
   '
   Exit Sub

90 XX$ = REGBLAN$("LIMAYOL")
   FFAX$ = Space$(8)
   If FECHASIEN% > 0 Then
      FEX = FECHASIEN%: FFAX$ = FECHATEX$(FEX)
   End If
   Call REPLA(XX$, FFAX$, 1, 8)
   If RENUMERA% = 1 Then
     NROASIEN = NROASDEF%
   End If
   Call REPLA(XX$, MKS$(NROASIEN), 9, 4)
   'Call REPLA(XX$, MKI$(NROASDEF%), 13, 2)
   Call REPLA(XX$, MARCADIA$, 15, 2)
   Call REPLA(XX$, REFE$, 17, 56)
   Call REPLA(XX$, MKD$(IDEBE#), 73, 8)
   Call REPLA(XX$, MKD$(IHABE#), 81, 8)
   SUTRA# = SUTRA# + IDEBE# - IHABE#
   Call REPLA(XX$, MKD$(SUTRA#), 89, 8)
   '
95 JJ& = JJ& + 1
   Call GRAREG("LIMAYOL", XX$, JJ&)
   Return
   '
End Sub
