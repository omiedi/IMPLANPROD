VERSION 5.00
Begin VB.Form RECICOMP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza - Opciones"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4095
      Picture         =   "RECICOMP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Control por Fecha de Vencimiento Cobranza"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Analisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4095
      Picture         =   "RECICOMP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Análisis de Cobranza por Fecha Cheques"
      Top             =   3360
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   26
      Top             =   2205
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4095
      Picture         =   "RECICOMP.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Listados de Recibos Registrados"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   4095
      Picture         =   "RECICOMP.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
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
      Height          =   750
      Left            =   4095
      Picture         =   "RECICOMP.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   3990
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   1
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   22
         Top             =   840
         Width           =   2430
         Begin VB.CommandButton Command6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Actualizar Tabla de Talonarios de Recibo y Cobradores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Cobradores / Talo- narios de Recibo"
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
            Left            =   735
            TabIndex        =   24
            Top             =   105
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Actualizar Tabla de Valores y Medios de Cobranza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tabla de Medios de Cobranza"
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
            Left            =   735
            TabIndex        =   10
            Top             =   0
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operaciones Especiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   19
         Top             =   2835
         Width           =   2430
         Begin VB.CommandButton Command5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":0E46
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ajuste Directo de Saldos Deudores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldos Iniciales en Cuentas Corrientes"
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
            Left            =   630
            TabIndex        =   21
            Top             =   75
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   16
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":0F90
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ajuste Directo de Saldos Deudores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Sal- dos Deudores"
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
            Left            =   840
            TabIndex        =   18
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1830
            Picture         =   "RECICOMP.frx":10DA
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Visualizar Imputación de Créditos y Anticipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":13E4
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Imputación de Créditos y Anticipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
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
            Left            =   720
            TabIndex        =   15
            Top             =   75
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":152E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Recibo ya Emitido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Repetir Impre- sion de Recibo"
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
            Left            =   840
            TabIndex        =   3
            Top             =   80
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "RECICOMP.frx":1B98
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de Recibo Emitido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Recibo Cobranza"
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
            Left            =   840
            TabIndex        =   6
            Top             =   80
            Width           =   1530
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3540
      Picture         =   "RECICOMP.frx":1EA2
      Top             =   5155
      Width           =   2010
   End
End
Attribute VB_Name = "RECICOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    Call CARTAMECO
    Command1.Enabled = True
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call COCALCOB
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    OPX% = COMALTER%("Puede Visualizar Aplicaciones de Diversas Formas:\\Por Recibo o N/C\Por Factura o N/D\Por Cliente")
    If OPX% = 1 Then
        Call VAPLIREC
      ElseIf OPX% = 2 Then
        Call VAPLIFAC
      ElseIf OPX% = 3 Then
        Call FINAL("*APLICOB")
    End If
    Command11.Enabled = True
End Sub

Sub VAPLIREC()
    '
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 RFF$ = RECFILT$("APLICOB", 2, MKI$(NC%))
    If Len(RFF$) < 4 Then
      Call MENSERR(24, "Sin Movimientos")
      GoTo 10
    End If
    '
    JJ& = 0
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      X$ = REGLEIDO("APLICOB", U&)
      NROX$ = Mid$(X$, 5, 18)
      IMPO# = CVD(Mid$(X$, 23, 8))
      If TRIM$(NROX$) <> "" Then
        For KU& = ULTREG&("!SELCUERE") To 1 Step -1
          YY$ = REGLEIDO$("!SELCUERE", KU&)
          If Mid$(YY$, 5, 18) = NROX$ Then
            IMPO# = IMPO# + Val(Mid$(YY$, 23, 13))
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(YY$, IMPOS$, 23, 13)
            Call GRAREG("!SELCUERE", YY$, KU&)
            GoTo 309
          End If
        Next KU&
        '
        Y$ = REGBLAN$("!SELCUERE")
        Call REPLA(Y$, MKS$(U&), 1, 4)
        Call REPLA(Y$, NROX$, 5, 18)
        IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
        Call REPLA(Y$, IMPOS$, 23, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELCUERE", Y$, JJ&)
      End If
309 Next UL&
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC%)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      JJ& = 0
      For UL& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UL&, 4))
        X$ = REGLEIDO("APLICOB", U&)
        NROX$ = Mid$(X$, 5, 18)
        If NROX$ = DAPLI$ Then
          NROY$ = Mid$(X$, 55, 18)
          REFY$ = Mid$(X$, 31, 24)
          IMPO# = CVD(Mid$(X$, 23, 8))
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          '
          Y$ = REGBLAN$("!SELDOCAP")
          Call REPLA(Y$, NROY$, 1, 18)
          Call REPLA(Y$, REFY$, 19, 24)
          Call REPLA(Y$, IMPOS$, 44, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELDOCAP", Y$, JJ&)
        End If
      Next UL&
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    GoTo 10
    '
999 End Sub
'
Sub VAPLIFAC()
    '
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 RFF$ = RECFILT$("APLICOB", 2, MKI$(NC%))
    If Len(RFF$) < 4 Then
      Call MENSERR(24, "Sin Movimientos")
      GoTo 10
    End If
    '
    JJ& = 0
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      X$ = REGLEIDO("APLICOB", U&)
      NROX$ = Mid$(X$, 55, 18)
      IMPO# = CVD(Mid$(X$, 23, 8))
      If TRIM$(NROX$) <> "" Then
        For KU& = ULTREG&("!SELCUERE") To 1 Step -1
          YY$ = REGLEIDO$("!SELCUERE", KU&)
          If Mid$(YY$, 5, 18) = NROX$ Then
            IMPO# = IMPO# + Val(Mid$(YY$, 23, 13))
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(YY$, IMPOS$, 23, 13)
            Call GRAREG("!SELCUERE", YY$, KU&)
            GoTo 309
          End If
        Next KU&
        '
        Y$ = REGBLAN$("!SELCUERE")
        Call REPLA(Y$, MKS$(U&), 1, 4)
        Call REPLA(Y$, NROX$, 5, 18)
        IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
        Call REPLA(Y$, IMPOS$, 23, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELCUERE", Y$, JJ&)
      End If
309 Next UL&
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC%)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      JJ& = 0
      For UL& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UL&, 4))
        X$ = REGLEIDO("APLICOB", U&)
        NROX$ = Mid$(X$, 55, 18)
        If NROX$ = DAPLI$ Then
          NROY$ = Mid$(X$, 5, 18)
          REFY$ = Mid$(X$, 31, 24)
          IMPO# = CVD(Mid$(X$, 23, 8))
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          '
          Y$ = REGBLAN$("!SELDOCAP")
          Call REPLA(Y$, NROY$, 1, 18)
          Call REPLA(Y$, REFY$, 19, 24)
          Call REPLA(Y$, IMPOS$, 44, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELDOCAP", Y$, JJ&)
        End If
      Next UL&
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    GoTo 10
    '
999 End Sub
'

Private Sub Command2_Click()
    '
    NROPRO& = 0
    Call FORECISTA.NURECIBO(NROPRO&)
    FORECISTA.Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    FORECISTA.Text11.SelStart = 3
    FORECISTA.Text11.SelLength = Len(FORECISTA.Text11.TEXT) - 3
    '
    Hide
    '
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
    If DERACCE%("REPIRECI", "Repetir Impresion de Recibo Emitido") >= 1 Then
       Call REPIRECI
    End If
    Command23.Enabled = True
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    'If DERACCE%("ANURECI", "Anulación de Recibos") >= 2 Then
       Call ANURECI
    'End If
    Command25.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'If DERACCE%("IMPUCRED", "Imputacion Créditos/Anticipos Deudores") >= 2 Then
      Hide
      Call FRESHALL
      Call IMPUCRE1
    'End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'If DERACCE%("AJUSALD", "Ajuste de Saldos Deudores Clientes") >= 2 Then
       Call AJUSALD
    'End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CARSALINID
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CARTACOBRA
    Command6.Enabled = True
End Sub


Private Sub Command7_Click()
    Call HELPONLINE("RECIBO")
End Sub

Sub CARSALINID()
   '
   If ULTREG&("SALINID") < 1 Then
      GoSub 4000
   End If
   '
2  X$ = REGLEIDO$("SALINID", 1)
   FEX = CVI(Mid$(X$, 9, 2))
   FSI% = FEX
   FAJU$ = FECHATEX$(FEX)
   '
3  CX% = ULTREG&("DEUDOR1")
   Dim SALINI#(32767), SUMDEU#(32767), SARRAS#(32767), REGAJU&(32767)
   '
   GoSub 6000    ' generar archivos eco de tipo y variante
   '
5  CLX$ = OBJPLA$("SELECLI", MKI$(NC%))
   If CLX$ = "" Then GoTo 200                   ' ajustar
   NC% = CVI(CLX$): If NC% < 1 Then Print Chr$(7);: GoTo 5
   CLI$ = RASOCLI$(NC%)
   '
   Call BLANARCH("!AJUSALD")
   '
6  Screen.MousePointer = 11
   FIN& = ULTREG&("SALINID")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("SALINID", I&)
      If CVI(Left$(X$, 2)) = NC% Then
          SIG% = 1
          TI% = Asc(Mid$(X$, 3, 1))
          VA% = Asc(Mid$(X$, 4, 1))
          If TI% < 0 Or TI% > 6 Then TI% = 2
          If TI% > 3 Then SIG% = (-1)
          Y$ = Mid$(X$, 3, 26) + MKS$(I&) + Mid$(X$, 31, 2)
          IPEN# = CVD(Mid$(X$, 11, 8)) * SIG%
          Call REPLA(Y$, MKD$(IPEN#), 9, 8)
          Call REGAPP("!AJUSALD", Y$)
      End If
   Next I&
   Screen.MousePointer = 1
   '
8  VTB% = VENTABU%("AJUSALD/NC/TITUPAN=" + CLI$)
   If VTB% < 0 Then GoTo 5
   ATX% = 0: ATY% = 0
   SALDEU# = 0
   '
   For U& = 1 To ULTREG&("!AJUSALD")
       '
       X$ = REGLEIDO$("!AJUSALD", U&)
       TI% = Asc(Left$(X$, 1))
       If TI% > 6 Then TI% = 2
       If TI% = 4 Then
          Call MENSERR(24, "Imposible Registrar Recibo\Como Saldo Inicial.\  \Registrar Como Ajuste de Saldo.")
          GoTo 8
       End If
       '
       If CVI(Mid$(X$, 17, 2)) > FSI% Then
          Call MENSERR(24, "Fecha de Origen Debe Ser Anterior\a Fecha Inicial.")
          GoTo 8
       End If
       '
       SIG% = 1: If TI% > 3 Then SIG% = (-1)
       SALEI# = CVD(Mid$(X$, 9, 8))
       SORI# = CVD(Mid$(X$, 19, 8))
       If Mid$(X$, 31, 2) <> String$(2, 255) Then
         If Abs(SALEI#) >= 0.005 Then
           If TI% > 0 Then
             If Sgn(SALEI#) <> SIG% Or Sgn(SORI#) <> Sgn(SALEI#) Then
               SALEI# = Abs(SALEI#) * SIG%
               Call REPLA(X$, MKD$(SALEI#), 9, 8)
               SORI# = (Abs(CVD(Mid$(X$, 19, 8)))) * Sgn(SALEI#)
               Call REPLA(X$, MKD$(SORI#), 19, 8)
               Call GRAREG("!AJUSALD", X$, U&)
               ATX% = 1
             End If
           End If
         End If
         ATY% = 1
       End If
       '
       SALDEU# = SALDEU# + SALEI#
   Next U&
   '
   If ATX% > 0 Then
       Call COMUNI("Errores de Signo de Saldos Deudores\han sido Corregidos Automáticamente.")
       GoTo 8
   End If
   '
100 FIN& = ULTREG&("!AJUSALD")
    Call BLOQARCH("SALINID")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("!AJUSALD", I&)
      REG& = CVS(Mid$(X$, 27, 4))
      '
      TI% = Asc(Left$(X$, 1))
      IPEN# = CVD(Mid$(X$, 9, 8)) * SIG%
      If TI% > 0 Then IPEN# = Abs(IPEN#)
      Y$ = MKI$(NC%) + Left$(X$, 26) + MKI$(0) + Mid$(X$, 31, 2)
      Call REPLA(Y$, MKD$(IPEN#), 11, 8)
      If REG& < 1 Then
           Call REGAPP("SALINID", Y$)
         Else
           If Mid$(Y$, 31, 2) <> String$(2, 255) Then
             Call GRAREG("SALINID", Y$, REG&)
           End If
      End If
   Next I&
   '
   GoTo 5
   '
   '
200 For CX1% = 1 To CX%
      SALINI#(CX1%) = 0
      SARRAS#(CX1%) = 0
      SUMDEU#(CX1%) = 0
    Next CX1%
    '
    Call BLOQARCH("SVD202")
    Call BLOQARCH("FACTUR")
    Call BLOQARCH("SALINID")
    FIN& = ULTREG&("SALINID")
    For I& = 2 To FIN&
      X$ = REGLEIDO$("SALINID", I&)
      If Len(X$) >= 32 Then
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           SIG% = 1: TI% = Asc(Mid$(X$, 3, 1))
           VA% = Asc(Mid$(X$, 4, 1))
           If TI% < 0 Or TI% > 6 Then TI% = 2
           If TI% > 3 Then SIG% = (-1)
           '
           IMPO# = CVD(Mid$(X$, 11, 8)) * SIG%
           'If TI% > 3 Then IMPO# = (-1) * Abs(IMPO#)
           '
           If Mid$(X$, 31, 2) <> String$(2, 255) Then
             NRODO& = CVS(Mid$(X$, 5, 4))
             Call REPLA(X$, String$(2, 255), 31, 2)
             Call GRAREG("SALINID", X$, I&)
             If TI% > 0 Then
                Call REPLA(X$, MKD$(Abs(IMPO#)), 21, 8)
             End If
             Call REGAPP("SVD202", X$)
             Y$ = REGBLAN$("FACTUR")
             If TI% > 0 Then IMPO# = Abs(IMPO#)
             Call REPLA(Y$, MKI$(NC%) + MKS$(NRODO&) + MKI$(FSI%) + MKI$(FSI%) + String$(2, 0) + Chr$(TI%) + Chr$(VA%), 1, 14)
             Call REPLA(Y$, MKD$(IMPO#), 15, 8)
             Call REPLA(Y$, MKD$(IMPO#), 39, 8)
             Call REGAPP("FACTUR", Y$)
           End If
         End If
      End If
    Next I&
    '
    Call CIERRARCH("SALINID")
    Call CIERRARCH("FACTUR")
    Call CIERRARCH("SVD202")
    '
    Exit Sub
    '
1000 '************************************************ listar saldos iniciales
   '
   Call BLOQARCH("LISALIN")
   Call BLANARCH("LISALIN")
   '
   FIN& = ULTREG&("SALINID")
   For I& = 2 To FIN&
    Call TRACE(20, I&, FIN&)
    X$ = REGLEIDO$("SALINID", I&)
    NC% = CVI(Left$(X$, 2))
    If NC% > 0 Then
      TI% = Asc(Mid$(X$, 3, 1))
      TIX$ = "FC": SIG% = 1
          If TI% = 3 Then TIX$ = "ND"
          If TI% = 4 Then TIX$ = "RB": SIG% = (-1)
          If TI% = 6 Then TIX$ = "NC": SIG% = (-1)
      NRO& = CVS(Mid$(X$, 5, 4))
      NROX$ = TIX$ + AJUSTD$(TRIM$(Str$(NRO&)), 8)
      FF% = CVI(Mid$(X$, 19, 2))
      IORI# = Abs(CVD(Mid$(X$, 21, 8)))
      FVI% = CVI(Mid$(X$, 9, 2))
      IPEN# = CVD(Mid$(X$, 11, 8)) * SIG%
      DEBE# = IPEN#: HABE# = 0
      If Sgn(IPEN#) < 0 Then DEBE# = 0: HABE# = Abs(IPEN#)
      '
      Z$ = REGBLAN$("LISALIN")
      Call REPLA(Z$, MKI$(NC%), 1, 2)
      Call REPLA(Z$, NROX$, 3, 10)
      Call REPLA(Z$, MKI$(FF%), 13, 2)
      Call REPLA(Z$, MKD$(IORI#), 15, 8)
      Call REPLA(Z$, MKI$(FVI%), 23, 2)
      Call REPLA(Z$, MKD$(DEBE#), 25, 8)
      Call REPLA(Z$, MKD$(HABE#), 33, 8)
      Call REPLA(Z$, MKD$(0), 41, 8)
      '
      Call REGAPP("LISALIN", Z$)
      '
    End If
    '
   Next I&
   '
   Call CIERRARCH("LISALIN")
   Call FINAL("*SOSALIN")
   '

4000 '************************************** ingresa fecha de saldos iniciales
   '
   HO% = HOY(HOS$)
   Call SELECHO("SELFECHA")
   FAJU$ = TRIM$(VALACT1$("SELFECHA"))
   If FAJU$ = "" Then Exit Sub
   '
   FF% = FECHANUM(FAJU$)
   If FF% > HO% Then GoTo 4000
   '
   FSI% = FF%
   X$ = REGBLAN$("SALINID")
   X$ = REPLACE$(X$, MKI$(FSI%), 9, 2)
   X$ = REPLACE$(X$, MKI$(FSI%), 19, 2)
   Call GRAREG("SALINID", X$, 1)
   Call CIERRARCH("SALINID")
Return
   '

6000 '  ******************************* genera archivos eco de tipo y variante
   '
   If ULTREG&("ECOTIPO") <> 7 Then
       Call BLANARCH("ECOTIPO")
       Call REGAPP("ECOTIPO", MKI$(0) + "Ajuste de Saldo")
       Call REGAPP("ECOTIPO", MKI$(1) + "Venta Contado")
       Call REGAPP("ECOTIPO", MKI$(2) + "Factura C.Corr.")
       Call REGAPP("ECOTIPO", MKI$(3) + "Nota de Debito")
       Call REGAPP("ECOTIPO", MKI$(4) + "Anticipo")
       Call REGAPP("ECOTIPO", MKI$(5) + "Descuento Caja")
       Call REGAPP("ECOTIPO", MKI$(6) + "Nota de Credito")
       Call CIERRARCH("ECOTIPO")
   End If
   '
   If ULTREG&("ECOVARI") <> 4 Then
       Call BLANARCH("ECOVARI")
       Call REGAPP("ECOVARI", MKI$(0) + "No Definido")
       Call REGAPP("ECOVARI", MKI$(1) + "A - Responsables IVA")
       Call REGAPP("ECOVARI", MKI$(2) + "B - Consumidor Final")
       Call REGAPP("ECOVARI", MKI$(3) + "X - Exportacion")
       Call CIERRARCH("ECOVARI")
   End If
   '
   Return

End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Dim RERECI&(10000)
   '
10 FEX = HOY(HO$)
   HAS% = FEX
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      GoTo 99
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   '
20 NUSUC% = 0
   If ULTREG&("COBRADOR") > 0 Then
      Call SELECHO("COBRADOR")
      NUSUC% = Val(TRIM$(VALACT1$("COBRADOR")))
      If NUSUC% < 1 Then GoTo 10
   End If
   '
   Screen.MousePointer = 11
   For JJ& = 0 To 10000: RERECI&(JJ&) = 0: Next JJ&
   PRIRECI& = 999999: ULTRECI& = 0
   '
   FIN& = ULTREG&("COBRAN")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("COBRAN", IL&)
     FF% = CVI(Mid$(X$, 7, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         If NUSUC% = 0 Or CVI(Mid$(X$, 23, 2)) = NUSUC% Then
           NRO& = CVS(Mid$(X$, 3, 4))
           If NRO& > 0 Then
             If NRO& < PRIRECI& Then PRIRECI& = NRO&
             If NRO& > ULTRECI& Then ULTRECI& = NRO&
           End If
         End If
       End If
     End If
   Next IL&
   Screen.MousePointer = 1
   '
   If ULTRECI& < PRIRECI& Then
     Call MENSERR(24, "Sin Movimientos\en el Rango de Fechas Elegido")
     GoTo 10
   End If
   '
   DEMARECI% = 0
   If ULTRECI& - PRIRECI& > 9998 Then
     DEMARECI% = 1
      Call MENSERR(24, "Demasiados Recibos.\  \Elija Rango de Fechas Menor.")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   REAVA& = 0
   FIN& = ULTREG&("COBRAN")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("COBRAN", IL&)
     FF% = CVI(Mid$(X$, 7, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         If NUSUC% = 0 Or CVI(Mid$(X$, 23, 2)) = NUSUC% Then
           NRO& = CVS(Mid$(X$, 3, 4))
           If NRO& > 0 Then
             If NRO& >= PRIRECI& Then
               If NRO& <= ULTRECI& Then
                 IJ& = NRO& - PRIRECI&
                 PFX$ = TRIM$(Str$(NUSUC%)): While Len(PFX$) < 2: PFX$ = "0" + PFX$: Wend
                 NNX$ = TRIM$(Str$(NRO&)): While Len(NNX$) < 6: NNX$ = "0" + NNX$: Wend
                 NNX$ = "RB-" + PFX$ + "-" + NNX$
                 If RERECI&(IJ&) > 0 Then
                     Call MENSERR(24, "Recibo " + NNX$ + " Repetido")
                   Else
                     RERECI&(IJ&) = IL&
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next IL&
   '
   FIN& = ULTRECI& - PRIRECI&
   For IL& = 0 To FIN&
     Call TRACE(20, IL&, FIN&)
     RERE& = RERECI&(IL&)
     NRO& = PRIRECI& + IL&
     PFX$ = TRIM$(Str$(NUSUC%)): While Len(PFX$) < 2: PFX$ = "0" + PFX$: Wend
     NNX$ = TRIM$(Str$(NRO&)): While Len(NNX$) < 6: NNX$ = "0" + NNX$: Wend
     NNX$ = "RB-" + PFX$ + "-" + NNX$
     Y$ = REGBLAN$("AVAFACB")
     Call REPLA(Y$, NNX$, 31, 18)
     '
     If RERE& > 0 Then
       If RERE& <= ULTREG&("COBRAN") Then
         X$ = REGLEIDO$("COBRAN", RERE&)
         NC% = CVI(Left$(X$, 2))
         FF% = CVI(Mid$(X$, 7, 2))
         NUSUCU% = CVI(Mid$(X$, 23, 2))
         '
         CLIE$ = RASOCLI$(NC%)
         NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100)
         '
         Call REPLA(Y$, MKI$(FF%), 1, 2)
         Call REPLA(Y$, CLIE$, 3, 28)
         Call REPLA(Y$, MKD$(NETO#), 49, 8)
       End If
     End If
     '
     REAVA& = REAVA& + 1
     Call GRAREG("AVAFACB", Y$, REAVA&)
   Next IL&
   '
   Call SETULTREG("AVAFACB", REAVA&)
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Call CONECRUN("*LIAVAFB", "Secuencia de Recibos Registrados/WAIT")
   If ULTREG&("COBRADOR") > 1 Then
     GoTo 20
   End If
   '
99 Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("ANARECIB", "Analisis de Cobranza por Cheque")
   Command9.Enabled = True
End Sub

Sub COCALCOB()
   '
   Dim NRECI&(8192)
   Dim NCLIE%(8192)
   Dim RECOB&(8192)
   Dim NUFAC&(8192)
   Dim TIFAC%(8192)
   '
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   REAVA& = 0: PRIRECI& = 9999999: ULTRECI& = 0
   CORRESP$ = MESANO$(MESASEL$, DES%, HAS%)
   FIN& = ULTREG&("COBRAN")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("COBRAN", IL&)
     FF% = CVI(Mid$(X$, 7, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100)
         If NETO# >= 0.005 Then
           NC% = CVI(Left$(X$, 2))  ' CLIE$ = RASOCLI$(NC%)
           NUSUCUVE% = CVI(Mid$(X$, 23, 2))
           NRO& = CVS(Mid$(X$, 3, 4))
           'PFX$ = TRIM$(Str$(NUSUCUVE%)): While Len(PFX$) < 2: PFX$ = "0" + PFX$: Wend
           'NNX$ = TRIM$(Str$(NRO&)): While Len(NNX$) < 6: NNX$ = "0" + NNX$: Wend
            NNX$ = "RB-" + PFX$ + "-" + NNX$
           '
           If NRO& < PRIRECI& Then PRIRECI& = NRO&
           If NRO& > ULTRECI& Then ULTRECI& = NRO&
           REAVA& = REAVA& + 1
           NRECI&(REAVA&) = NRO&
           NCLIE%(REAVA&) = NC%
           RECOB&(REAVA&) = IL&
         End If
       End If
     End If
   Next IL&
   '
   FIN& = ULTREG&("APLICOB")
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     X$ = REGLEIDO$("APLICOB", IL&)
     DOCAP$ = Mid$(X$, 5, 18)
     If Mid$(DOCAP$, 17, 1) = "-" Then
       DOCEP$ = Left$(DOCAP$, 2) + Mid$(DOCAP$, 17, 2) + Mid$(DOCAP$, 3, 14)
       Call REPLA(X$, DOCEP$, 5, 18)
       Call GRAREG("APLICOB", X$, IL&)
     End If
     If Mid$(X$, 5, 2) = "RB" Then
       NRO& = Val(Mid$(X$, 15, 8))
       If NRO& >= PRIRECI& Then
         If NRO <= ULTRECI& Then
           IAPLI# = CVD(Mid$(X$, 23, 8))
           If IAPLI# >= 0.005 Then
             DOCOR$ = Mid$(X$, 55, 18)
             TIDOC% = 2
             If Left$(DOCOR$, 2) = "ND" Then TIDOC% = 3
             NC% = CVI(Mid$(X$, 3, 2))
             '
             For JJ& = 1 To REAVA&
               If NRECI&(JJ&) = NRO& Then
                 If NCLIE%(JJ&) = NC% Then
                   If NUFAC&(JJ&) < 1 Then
                       NUFAC&(JJ&) = Val(Mid$(DOCOR$, 9, 8))
                       TIFAC%(JJ&) = TIDOC%
                     Else
                       REAVA& = REAVA& + 1
                       NRECI&(REAVA&) = NRO&
                       NCLIE%(REAVA&) = NC%
                       RECOB&(REAVA&) = RECOB&(JJ&)
                       NUFAC&(REAVA&) = Val(Mid$(DOCOR$, 9, 8))
                       TIFAC%(REAVA&) = TIDOC%
                   End If
                   GoTo 19
                 End If
               End If
             Next JJ&
             '
           End If
         End If
       End If
     End If
19 Next IL&
   '
   REBLA$ = REGBLAN$("CTRLCOB")
   JJ& = 0
   FIN& = REAVA&
   For IL& = 1 To FIN&
     Call TRACE(20, IL&, FIN&)
     NC% = NCLIE%(IL&)
     RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
     For KKL& = 1 To Len(RFF$) Step 4
       REFAC& = CVS(Mid$(RFF$, KKL&, 4))
       Y$ = REGLEIDO$("FACTUR", REFAC&)
       If CVS(Mid$(Y$, 3, 4)) = NUFAC&(IL&) Then
         If Asc(Mid$(Y$, 13, 1)) = TIFAC%(IL&) Then
           '
           ZZ$ = REBLA$
           X$ = REGLEIDO$("COBRAN", RECOB&(IL&))
           FERECI% = CVI(Mid$(X$, 7, 2))
           NURECI& = CVS(Mid$(X$, 3, 4))
           IMRECI# = CVD(Mid$(X$, 15, 8))
           FEFAC% = CVI(Mid$(Y$, 7, 2))
           CPAG% = CPAGCLI%(NC%)
           REBU& = REGBUS&("CONPAG", 1, Chr$(CPAG%))
           PLA1% = 0
           If REBU& > 0 Then
             CPX$ = REGLEIDO$("CONPAG", REBU&)
             PLA1% = Asc(Mid$(CPX$, 38, 1))
           End If
           FEFA = FEFAC%: FEBE = FERECI%
           DIREA% = DIENTRE%(FEFA, FEBE)
           ATRAX% = DIREA% - PLA1%
           If ATRAX% < 0 Then ATRAX% = 0
           
           Call REPLA(ZZ$, MKI$(FERECI%), 1, 2)
           Call REPLA(ZZ$, MKS$(NURECI&), 3, 4)
           Call REPLA(ZZ$, MKD$(IMRECI#), 7, 8)
           Call REPLA(ZZ$, MKI$(NC%), 15, 2)
           Call REPLA(ZZ$, MKI$(TIFAC%(IL&)), 17, 2)
           Call REPLA(ZZ$, MKS$(NUFAC&(IL&)), 19, 4)
           Call REPLA(ZZ$, MKI$(FEFAC%), 23, 2)
           Call REPLA(ZZ$, MKI$(CPAG%), 25, 2)
           Call REPLA(ZZ$, MKI$(PLA1%), 27, 2)
           Call REPLA(ZZ$, MKI$(DIREA%), 29, 2)
           Call REPLA(ZZ$, MKI$(ATRAX%), 31, 2)
           GoTo 29
           '
         End If
       End If
     Next KKL&
     '
     ZZ$ = REBLA$
     X$ = REGLEIDO$("COBRAN", RECOB&(IL&))
     FERECI% = CVI(Mid$(X$, 7, 2))
     NURECI& = CVS(Mid$(X$, 3, 4))
     IMRECI# = CVD(Mid$(X$, 15, 8))
     '
     Call REPLA(ZZ$, MKI$(FERECI%), 1, 2)
     Call REPLA(ZZ$, MKS$(NURECI&), 3, 4)
     Call REPLA(ZZ$, MKD$(IMRECI#), 7, 8)
     Call REPLA(ZZ$, MKI$(NC%), 15, 2)
     '
29   JJ& = JJ& + 1
     Call GRAREG("CTRLCOB", ZZ$, JJ&)
     '
   Next IL&
   Call CIERRARCH("CTRLCOB")
   '
   Call FILESORT("CTRLCOB", "", 1, 4, "ASC")
   Call FINAL("*CTRLCOB")
   '
End Sub
