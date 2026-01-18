VERSION 5.00
Begin VB.Form RECICO00 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza - Opciones"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   19
      Top             =   3780
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   20
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
      Height          =   855
      Left            =   4095
      Picture         =   "RECICO00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Listados de Recibos Registrados"
      Top             =   2100
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
      Height          =   855
      Left            =   4095
      Picture         =   "RECICO00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
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
      Height          =   855
      Left            =   4095
      Picture         =   "RECICO00.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1220
      Width           =   855
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
      Height          =   4350
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   12
         Top             =   2520
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
            Picture         =   "RECICO00.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   14
            Top             =   75
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   9
         Top             =   1995
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
            Left            =   1840
            Picture         =   "RECICO00.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   24
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
            Picture         =   "RECICO00.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Left            =   735
            TabIndex        =   11
            Top             =   75
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   1470
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
            Picture         =   "RECICO00.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Recibo ya Emitido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver por Pantalla y Re-Imprimir"
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
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   21
         Top             =   945
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
            Picture         =   "RECICO00.frx":1006
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Imprimir Copia de Orden de Pago ya Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Des-Aplicar de Recibos de Cob."
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
            TabIndex        =   23
            Top             =   75
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   15
         Top             =   3465
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
            Picture         =   "RECICO00.frx":1150
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   17
            Top             =   75
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
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
            Picture         =   "RECICO00.frx":129A
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
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   5145
      Y1              =   4530
      Y2              =   4530
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3540
      Picture         =   "RECICO00.frx":15A4
      Top             =   4025
      Width           =   2010
   End
End
Attribute VB_Name = "RECICO00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

 Sub Command23_Click()
    '
    Command23.Enabled = False
    If DERACCE%("REPIRECI", "Repetir Impresion de Recibo Emitido") >= 1 Then
       Call REPIRECI
    End If
    Command23.Enabled = True
    '
End Sub

 Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("ANURECI", "Anulación de Recibos") >= 2 Then
       Call ANURECIBO
    End If
    Command25.Enabled = True
End Sub

 Sub Command3_Click()
    Command3.Enabled = False
    If DERACCE%("IMPUCRED", "Imputacion Créditos/Anticipos Deudores") >= 2 Then
      Hide
      Call FRESHALL
      Call IMPUCRE1
    End If
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AJUSALD", "Ajuste de Saldos Deudores Clientes") >= 2 Then
       Call AJUSALCLI
    End If
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
   Command5.Enabled = False
   Call COMUNI("Opción no Disponible en la Presente Versión.\  \Carga de Saldos Iniciales Clientes\por Ajuste de Saldos Deudores.")
   'Call CARSALINID
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call COMUNI("No Disponible")
   'Call DESAPLIRECI
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

 Sub Command8_Click()
   '
   Command8.Enabled = False
   MODO% = ALTERNA%("Avance de Cobranza\Valores por Comprobante\Detalle por Tipo de Valor")
   Select Case MODO%
          Case 1
                Dim RERECI&(10000)
                '
10              FEX = HOY(HO$)
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
20              NUSUC% = 0
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
          Case 2, 3
                I& = 0
                FEX = HOY(HO$)
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
                Screen.MousePointer = 11
                FIN& = ULTREG&("CAJABANC")
                For IL& = 1 To FIN&
                  Call TRACE(20, IL&, FIN&)
                  X$ = REGLEIDO$("CAJABANC", IL&)
                  If Asc(Mid$(X$, 121, 1)) <> 9 Then ' Documento Anulado
                    FF% = CVI(Mid$(X$, 61, 2))
                    If FF% >= DES% Then
                      If FF% <= HAS% Then
                        If Left$(X$, 2) = "RB" Then
                          I& = I& + 1
                          Z$ = REGBLAN$("LISRECIB")
                          Call REPLA(Z$, Mid$(X$, 61, 2), 1, 2) ' Fecha de recibo
                          Call REPLA(Z$, Mid$(X$, 1, 12), 3, 12) ' Nro de recibo
                            NCLIE% = CVI(Mid$(X$, 67, 2)) ' Numero de cliente
                          Call REPLA(Z$, RASOCLI$(NCLIE%), 15, 28) ' Raz.Soc. Cliente
                            CODCUEN$ = MKI$(Val(CVI(Mid$(X$, 65, 2)))) ' Cod. interno cuenta contable
                            ZZ$ = FILTERGETRECORD$("ECUECON", 3, CODCUEN$)
                            CUENCONT$ = Mid$(ZZ$, 1, 12)
                          Call REPLA(Z$, CUENCONT$, 43, 12) ' Cuenta Contable
                          Call REPLA(Z$, Mid$(X$, 71, 40), 55, 40) ' Detalle del valor
                          Call REPLA(Z$, Mid$(X$, 63, 2), 95, 2) ' Fecha Acreditacion del valor
                          Call REPLA(Z$, Mid$(X$, 153, 8), 97, 8) ' Importe debe
                          Call REPLA(Z$, Mid$(X$, 161, 8), 105, 8) ' Importe haber
                          Call GRAREG("LISRECIB", Z$, I&)
                          Call SETULTREG("LISRECIB", I&)
                        End If
                      End If
                    End If
                  End If
                Next IL&
                '
                Screen.MousePointer = 1
                If I& > 0 Then
                  If MODO% = 2 Then
                    Call FILESORT("LISRECIB", "", 1, 2, "ASC") ' Ordena por Fecha y comprobante
                    Call HEADERS("LISRECIB", "")
                    Call HEADERS("LISRECIB", "Periodo: " + FECHATEX$(DES%) + " - " + FECHATEX(HAS%))
                    Call FINAL("*LISREC1")
                  Else
                    N& = 0
                    FIN& = ULTREG&("LISRECIB")
                    For U& = 1 To FIN&
                      X$ = REGLEIDO$("LISRECIB", U&)
                      If CVD(Mid$(X$, 97, 8)) <> 0 Then
                        N& = N& + 1
                        Z$ = Mid$(X$, 43, 12) & X$
                        Call GRAREG("LISREC2", Z$, N&)
                      End If
                    Next U&
                    Call SETULTREG("LISREC2", N&)
                    Call FILESORT("LISREC2", "", 1, 1, "ASC") ' Ordena por cuenta contable
                    Call HEADERS("LISREC2", "")
                    Call HEADERS("LISREC2", "Periodo: " + FECHATEX$(DES%) + " - " + FECHATEX(HAS%))
                    Call FINAL("*LISREC2")
                  End If
                End If
                Call CIERRARCH("*.*")
                '
   End Select
   '
99 Command8.Enabled = True
   '
End Sub

Sub ANURECIBO()
    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
10  Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    '
    Screen.MousePointer = 11
    FENUE$ = CONTROL$("", "FENURECI")
    FLIXX% = FECHANUM(FENUE$)
    If FLIXX% <= FECHALIM% Then FLIXX% = FECHALIM% + 1
    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 For U& = 1 To ULTREG&("COBRAN")
      X$ = REGLEIDO("COBRAN", U&)
      If CVI(Left$(X$, 2)) = NC% Then
        NROLL& = CVS(Mid$(X$, 3, 4))
        IMPO# = CVD(Mid$(X$, 15, 8))
        'If Abs(IMPO#) >= 0.005 Then
          FC% = CVI(Mid$(X$, 7, 2))
          If FC% >= FLIXX% Then
            If NROLL& > 0 Then
              IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
              NNXX# = CDbl(NROLL&)
              DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
              Y$ = REGBLAN$("!SELCUERE")
              Call REPLA(Y$, MKS$(U&), 1, 4)
              Call REPLA(Y$, DOCU$, 5, 18)
              Call REPLA(Y$, IMPOS$, 23, 13)
              Call REGAPP("!SELCUERE", Y$)
            End If
          End If
        'End If
      End If
    Next U&
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Screen.MousePointer = 1
    Call SELECHO("!SELCUERE/Selección Recibos para Anulación")
    REGCUEC& = Val(VALACT1$("!SELCUERE"))
    If Val(REGCUEC&) < 1 Then
        GoTo 999
    End If
    '
    XX0$ = REGLEIDO$("COBRAN", REGCUEC&)
    FERECIBO% = CVI(Mid$(XX0$, 7, 2))
    NUTALONA% = CVI(Mid$(XX0$, 23, 2))
    IMPO# = Abs(CVD(Mid$(XX0$, 15, 8)))
    IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
    NROLL& = CVS(Mid$(XX0$, 3, 4)): NNXX# = CDbl(NROLL&)
    X$ = "Recibo Nro" + NUMSTRI$(NNXX#, 8, 0, 1)
    Y$ = IMPOS$
    '
    If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Abandona\Confirma Anulacion") <> 2 Then GoTo 999
    '
    ' anulacion propiamente dicha
    ' ***************************
    '
311 Screen.MousePointer = 11
    Call BLOQARCH("COBRAN")
    Call BLOQARCH("SVD202")
    Call BLOQARCH("CAJABANC")
    Call BLOQARCH("APLICOB")
    '
    RFF$ = RECFILT$("CAJABANC", 7, MKI$(NC%))
    For URX& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, URX&, 4))
      Z$ = REGLEIDO$("CAJABANC", U&)
      If UCase$(Left$(Z$, 2)) = "RB" Then
        If Val(Mid$(Z$, 4, 6)) = NROLL& Then
          If Asc(Mid$(Z$, 121, 1)) = 2 Then
            Screen.MousePointer = 1
            Call CIERRARCH("COBRAN")
            Call CIERRARCH("SVD202")
            Call CIERRARCH("CAJABANC")
            Call CIERRARCH("APLICOB")
            Call COMUNI("Anulación Imposible.\ \Valores Ingresados con este Recibo\Figuran Depositados o Entregados en Pago.")
            GoTo 10
          End If
        End If
      End If
    Next URX&
    '
    Call REPLA(XX0$, MKD$(0), 15, 8)
    Call GRAREG("COBRAN", XX0$, REGCUEC&)
    Call CIERRARCH("COBRAN")
    '
    NC1% = NC%
    RC& = REGICLI&(NC1%)
    DX$ = REGLEIDO$("DEUDOR2", RC&)
    SALDO# = CVD(Mid$(DX$, 109, 8)) + IMPO#
    Call REPLA(DX$, MKD$(SALDO#), 109, 8)
    Call GRAREG("DEUDOR2", DX$, RC&)
    Call CIERRARCH("DEUDOR2")
    '
    RFF$ = RECFILT$("CAJABANC", 7, MKI$(NC%))
    For URX& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, URX&, 4))
      Z$ = REGLEIDO$("CAJABANC", U&)
      If UCase$(Left$(Z$, 2)) = "RB" Then
        If Val(Mid$(Z$, 4, 6)) = NROLL& Then
          If Asc(Mid$(Z$, 121, 1)) <> 9 Then
            Call REPLA(Z$, Chr$(9), 121, 1)
            Call REPLA(Z$, IDACTU$, 269, 6)
            Call GRAREG("CAJABANC", Z$, U&)
          End If
        End If
      End If
312 Next URX&
    Call CIERRARCH("CAJABANC")
    '
    J& = 0
    IAJU# = IMPO#
    For U& = ULTREG&("APLICOB") To 1 Step -1
        Z$ = REGLEIDO$("APLICOB", U&)
        If Mid$(Z$, 5, 2) = "RB" Then
          If Val(Mid$(Z$, 15, 8)) = NROLL& Then
            FFA% = CVI(Left$(Z$, 2))
            NCA% = CVI(Mid$(Z$, 3, 2))
            ICA# = CVD(Mid$(Z$, 23, 8))
            dca$ = Mid$(Z$, 55, 18)
            TIA% = 2
            TIAS$ = Left$(dca$, 2)
            If TIAS$ = "ND" Then TIA% = 3
            If TIAS$ = "AS" Or TIAS$ = "AJ" Then TIA% = 0
            NRA& = Val(Mid$(dca$, 9, 8))
            VAA% = 1
            VAAS$ = Mid$(dca$, 18, 1)
            If VAAS$ = "B" Then VAA% = 2
            If VAAS$ = "C" Then VAA% = 3
            ido# = CVD(Mid$(Z$, 73, 8))
            '
            Call REPLA(Z$, MKD$(0), 23, 8)
            Call GRAREG("APLICOB", Z$, U&)
            '
            If ICA# > IAJU# Then ICA# = IAJU#
            
            ZZ$ = REGBLAN$("SVD202")
            Call REPLA(ZZ$, MKI$(NCA%), 1, 2)
            Call REPLA(ZZ$, Chr$(TIA%), 3, 1)
            Call REPLA(ZZ$, Chr$(VAA%), 4, 1)
            Call REPLA(ZZ$, MKS$(NRA&), 5, 4)
            Call REPLA(ZZ$, MKI$(FFA%), 9, 2)
            Call REPLA(ZZ$, MKD$(ICA#), 11, 8)
            Call REPLA(ZZ$, MKI$(FFA%), 19, 2)
            Call REPLA(ZZ$, MKD$(ido#), 21, 8)
            Call REPLA(ZZ$, String$(2, 255), 31, 2)
            Call REGAPP("SVD202", ZZ$)
            IAJU# = IAJU# - ICA#
            If IAJU# < 0.005 Then GoTo 315
          End If
        End If
    Next U&
    '
315 If Abs(IAJU#) >= 0.005 Then
        Call COMUNI("No se han Encontrado Todos los Documentos\Cancelados por este Recibo " + TRIM$(Str$(NROLL&)) + ".\  \La Diferencia Quedó Registrada\como Ajuste de Saldo.")
        '
        Z$ = REGBLAN$("SVD202")
        Call REPLA(ZZ$, MKI$(NC%), 1, 2)
        Call REPLA(ZZ$, Chr$(0), 3, 1)
        Call REPLA(ZZ$, Chr$(0), 4, 1)
        Call REPLA(ZZ$, MKS$(NROLL&), 5, 4)
        Call REPLA(ZZ$, MKI$(FERECIBO%), 9, 2)
        Call REPLA(ZZ$, MKD$(IAJU#), 11, 8)
        Call REPLA(ZZ$, MKI$(FERECIBO%), 19, 2)
        Call REPLA(ZZ$, MKD$(IMPO#), 21, 8)
        Call REPLA(ZZ$, String$(2, 255), 31, 2)
        Call REGAPP("SVD202", ZZ$)
        '
    End If
    '
    Call CIERRARCH("SVD202")
    Call CIERRARCH("APLICOB")
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("COBRAN")
    Call LIBARCH("SVD202")
    Call LIBARCH("CAJABANC")
    Call LIBARCH("APLICOB")
    Call COMUNI("Anulación Recibo " + TRIM$(Str$(NROLL&)) + " Completa")
    RECUEC00.Text1.TEXT = TRIM$(Str$(NC%))
    Hide
    Call FRESHALL
    PROLLAMA$ = Name
    RECUEC00.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
End Sub
'

Sub AJUSALCLI()
    '
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%)
    CLSS$ = OBJPLA$("SELCLIAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$(NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    FIN& = Len(RFF$)
    NUSAL& = ULTREG&("FACTUR")
    For I1& = 1 To FIN& Step 4
       Call TRACE(20, I1&, FIN&)
       I& = CVS(Mid$(RFF$, I1&, 4))
       F0$ = REGLEIDO$("FACTUR", I&)
       F1$ = Left$(F0$, 2)
       '
       If CVI(F1$) = NC% Then
          '
          F3$ = Mid$(F0$, 7, 2)
          F6$ = Mid$(F0$, 13, 2)
          F7$ = Mid$(F0$, 15, 8)
          F8$ = Mid$(F0$, 23, 8)
          F9$ = Mid$(F0$, 31, 8)
          F12$ = Mid$(F0$, 49, 8)
          F13$ = Mid$(F0$, 57, 8)
          '
          FF% = CVI(F3$)
          '
          IMPO# = CVD(F7$) + CVD(F8$) + CVD(F9$) + CVD(F12$) + CVD(F13$)
          SIG% = 1: If IMPO# <= (-0.005) Then SIG% = (-1)
          IMPO# = Abs(IMPO#)
          TI% = Asc(F6$): If TI% < 0 Or TI% > 6 Then TI% = 2
          If TI% >= 4 Then If TI% <= 6 Then SIG% = -1
          '
          SALFINAL# = SALFINAL# + IMPO# * SIG%
          If FF% <= FAJUI% Then
              SARRASFE# = SARRASFE# + IMPO# * SIG%
          End If
          '
       End If
    Next I1&
    '
    RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    FIN& = Len(RFF$)
    NUSAL& = ULTREG&("COBRAN")
    For I1& = 1 To FIN& Step 4
       Call TRACE(20, I1&, FIN&)
       I& = CVS(Mid$(RFF$, I1&, 4))
       C0$ = REGLEIDO$("COBRAN", I&)
       C1$ = Left$(C0$, 2)
       '
       If CVI(C1$) = NC% Then
          '
          C3$ = Mid$(C0$, 7, 2)
          C5$ = Mid$(C0$, 15, 8)
          '
          FF% = CVI(C3$)
          IMPO# = CVD(C5$)
          '
          SALFINAL# = SALFINAL# - IMPO#
          If FF% <= FAJUI% Then
            SARRASFE# = SARRASFE# - IMPO#
          End If
          '
       End If
    Next I1&
    '
    RFF$ = RECFILT$("SVD202", 1, MKI$(NC%))
    FIN& = Len(RFF$)
    NUSAL& = ULTREG&("SVD202")
    For I1& = 1 To FIN& Step 4
       Call TRACE(20, I1&, FIN&)
       I& = CVS(Mid$(RFF$, I1&, 4))
       SS0$ = REGLEIDO$("SVD202", I&)
       SS1$ = Left$(SS0$, 2)
       If CVI(SS1$) = NC% Then
          '
          SS2$ = Mid$(SS0$, 3, 2)
          SS4$ = Mid$(SS0$, 9, 2)
          SS5$ = Mid$(SS0$, 11, 8)
          '
          FVIJ% = CVI(SS4$)
          If FVIJ% > 0 Then
             TI% = Asc(Mid$(SS2$, 1, 1))
             If TI% > 6 Then TI% = 2
             SALDEU# = CVD(SS5$)
             If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
             '
             If Abs(SALDEU#) >= 0.005 Then
                SALDEUTO# = SALDEUTO# + SALDEU#
             End If
          End If
          '
       End If
    Next I1&
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(FAJUI%) + Space$(256)
    '
    Screen.MousePointer = 1
    '
102 CLTT$ = OBJPLA$("CARAJUCLI", VDEF$)
    If CLTT$ = "" Or CLTT$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLTT$, 13, 8))
    FEVEN% = CVI(Mid$(CLTT$, 21, 2))
    If FEVEN% < FAJUI% Then GoTo 100
    '
    MOTIVO$ = Mid$(CLTT$, 23)
    If TRIM$(MOTIVO$) = "" Then
       Call MENSERR(24, "Debe Indicar un Motivo del Ajuste")
       GoTo 100
    End If
    '
    ICONI% = CVI(Mid$(CLTT$, 279, 2))
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
      Call MENSERR(24, "Falta Imputación Contable Ajuste")
      VDEF$ = CLTT$
      GoTo 102
    End If
    If ICONI% = CUECOCLI%((-1) * NC%) Then
      Call MENSERR(24, "Cuenta Contable no Válida")
      VDEF$ = CLTT$
      GoTo 102
    End If
    '
    DIFESAL# = SALAJU# - SARRASFE#
    DIFEDEU# = SALFINAL# - SALDEUTO# + DIFESAL#
    '
    If DIFESAL# < 0 Then
       FEVEN% = FAJUI%
    End If
    '
    Call CIERRARCH("FACTUR")
    Call CIERRARCH("SVD202")
    Call CIERRARCH("DEUDOR2")
    Call CIERRARCH("CAJABANC")
    '
    Call BLOQARCH("FACTUR")
    Call BLOQARCH("SVD202")
    Call BLOQARCH("DEUDOR2")
    Call BLOQARCH("CAJABANC")
    '
    SALFINAL# = SALFINAL# + DIFESAL#
    RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
        If RGX& <= ULTREG&("DEUDOR2") Then
           DX$ = REGLEIDO$("DEUDOR2", RGX&)
           DX$ = REPLACE$(DX$, MKD$(SALFINAL#), 109, 8)
           Call GRAREG("DEUDOR2", DX$, RGX&)
        End If
    End If
    '
    Call CIERRARCH("DEUDOR2")
    '
    HAYAJU% = 0
    If Abs(DIFESAL#) >= 0.005 Or Abs(DIFEDEU#) >= 0.005 Then
       '
       Screen.MousePointer = 11
       '
       FF0$ = REGBLAN$("FACTUR")
       NUAJU& = 1 + ULTREG&("FACTUR")
       Call REPLA(FF0$, MKI$(NC%), 1, 2)
       Call REPLA(FF0$, MKS$(NUAJU&), 3, 4)
       Call REPLA(FF0$, MKI$(FAJUI%), 7, 2)
       Call REPLA(FF0$, MKI$(FAJUI%), 9, 2)
       Call REPLA(FF0$, String$(4, 0), 11, 4)
       Call REPLA(FF0$, MKD$(DIFESAL#), 15, 8)
       Call REPLA(FF0$, MKD$(DIFESAL#), 39, 8)
       Call REGAPP("FACTUR", FF0$)
       '
       SS0$ = REGBLAN$("SVD202")
       Call REPLA(SS0$, MKI$(NC%), 1, 2)
       Call REPLA(SS0$, String$(2, 0), 3, 2)
       Call REPLA(SS0$, MKS$(NUAJU&), 5, 4)
       Call REPLA(SS0$, MKI$(FEVEN%), 9, 2)
       Call REPLA(SS0$, MKD$(DIFEDEU#), 11, 8)
       Call REPLA(SS0$, MKI$(FAJUI%), 19, 2)
       Call REPLA(SS0$, MKD$(DIFESAL#), 21, 8)
       Call REGAPP("SVD202", SS0$)
       '
       Call CIERRARCH("FACTUR")
       Call CIERRARCH("SVD202")
       '
       HAYAJU% = 1
       '
    End If
    '
    If InStr(UCase$(App.EXEName), "GES") < 1 Then
      GoTo 200
    End If
    '
    ' grabar AJUSTE en CAJABANC
    URECOBRA& = ULTREG&("FACTUR")
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    ORMA% = (-1)
    RFF$ = RECFILT$("CAJABANC", 3, MKI$(FAJUI%))
    For I& = 1 To Len(RFF$) Step 4
      RECAJ& = CVS(Mid$(RFF$, I&, 4))
      XY$ = REGLEIDO$("CAJABANC", RECAJ&)
      If Left$(XY$, 10) = NRECIB$ Then
        ORMB% = Val(Mid$(XY$, 11, 2))
        If ORMB% > ORMA% Then ORMA% = ORMB%
      End If
    Next I&
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    Z$ = REGBLAN$("CAJABANC")
    Call REPLA(Z$, NRECIB$, 1, 12)
    Call REPLA(Z$, RASOCLI$(NC%), 13, 48)
    Call REPLA(Z$, MKI$(FAJUI%), 61, 2)
    Call REPLA(Z$, MKI$(FAJUI%), 63, 2)
    Call REPLA(Z$, MKI$(CUECOCLI%(NC%)), 65, 2)
    Call REPLA(Z$, MKI$(NC%), 67, 2)
    Call REPLA(Z$, "AS", 69, 2)
    Call REPLA(Z$, "Ajuste Saldo " + NRECIB$, 71, 40)
    Call REPLA(Z$, MKI$(FAJUI%), 111, 2)
    Call REPLA(Z$, MKI$(0), 115, 2)
    Call REPLA(Z$, MKS$(URECOBRA&), 117, 4)
    Call REPLA(Z$, Chr$(0), 121, 1)
    Call REPLA(Z$, MKD$(DIFESAL#), 145, 8)
    If DIFESAL# < 0 Then
        Call REPLA(Z$, MKD$(Abs(DIFESAL#)), 161, 8)
      Else
        Call REPLA(Z$, MKD$(Abs(DIFESAL#)), 153, 8)
    End If
    Call REPLA(Z$, RASOCLI$((-1) * NC%), 177, 32)
    Call REPLA(Z$, "FACTUR", 245, 8)
    Call REPLA(Z$, MKS$(URECOBRA&), 253, 4)
    Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
    Call REGAPP("CAJABANC", Z$)
    '
    Call REPLA(Z$, "AJ.Saldo " + RASOCLI$(NC%), 13, 48)
    Call REPLA(Z$, MKI$(ICONI%), 65, 2)
    Call REPLA(Z$, String$(16, 0), 153, 16)
    If DIFESAL# < 0 Then
        Call REPLA(Z$, MKD$(Abs(DIFESAL#)), 153, 8)
      Else
        Call REPLA(Z$, MKD$(Abs(DIFESAL#)), 161, 8)
    End If
    Call REPLA(Z$, "Aj.Saldo " + RASOCLI$(NC%), 177, 32)
    Call REGAPP("CAJABANC", Z$)
    Call CIERRARCH("CAJABANC")
    '
200 Call CIERRARCH("*.*")
    Call LIBARCH("FACTUR")
    Call LIBARCH("SVD202")
    '
    If HAYAJU% > 0 Then
      Call GRAREG("!OBSERVOF", Left$(MOTIVO$, 64), 1)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 65, 64), 2)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 129, 64), 3)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 193, 64), 4)
      Call SETULTREG("!OBSERVOF", 4)
      Call CIERRARCH("!OBSERVOF")
      '
      FORIPRE$ = TRIM$(CONTROL$("", "FORAJUSA"))
      If TRIM$(FORIPRE$) <> "" Then
         '
         FEX = HOY(HOS$)
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = TRIM$(Str$(NUAJU&))
         TIPODOC$ = "Ajuste Saldo Cliente"
         DESTIDOC% = NC%
         '
         FEX = FAJUI%
         CODPARAM$(1) = "FECH-VEN": DOCPARAM$(1) = FECHATEX$(FEX)
         CODPARAM$(2) = "IMPO-PEN": DOCPARAM$(2) = CSTRING$(MKD$(SARRASFE#), 4, 12, 2, 2)
         CODPARAM$(3) = "IMPO-CAN": DOCPARAM$(3) = CSTRING$(MKD$(DIFESAL#), 4, 12, 2, 2)
         CODPARAM$(4) = "IMP-TOTA": DOCPARAM$(4) = CSTRING$(MKD$(SALAJU#), 4, 12, 2, 2)
         CAPARDOC% = 4
         CACOLTAB1% = 0
         CAITAB1% = 0
         '
         Call PRINTDOC("FORAJUSA")   ' Remito de Facturacion
         '
      End If
      Call SETULTREG("!OBSERVOF", 0)
      Call CIERRARCH("!OBSERVOF")
      '
      Screen.MousePointer = 1
      Call COMUNI("Ajuste Completo")
      '
    End If
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'

Sub DESAPLIRECI()       ' DES-APLICACION de Recibos
     '
     ' 1.- Ingresar Numero de Cliente
     '     **************************
     '
     FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
     '
10   Call SELECHO("DEUDOR1")
     NC% = Val(VALACT1$("DEUDOR1"))
     If NC% < 1 Then Exit Sub
     '
100  Call BLANARCH("!SELCUECO")
     MAXDEU& = ULTREG&("COBRAN")
     RFD$ = RECFILT$("COBRAN", 1, MKI$(NC%))
     '
     For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Mid$(X$, 1, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        '
        IMPO# = CVD(Mid$(X$, 15, 8))
        If IMPO# >= 0.005 Then
          IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
          ' Falta contemplar que pasa cuando hay varias talonarios
          COMPB$ = String$(18, " ")
          'COMPB$ = "RC-A-0000-" & MKI$(X$, 3, 4)
          Y$ = MKS$(REG&) + COMPB$ + IMPOS$
          Call REGAPP("!SELCUECO", Y$)
        End If
     Next U&
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Recibos de Cobranza que Puedan Desaplicarse.")
        GoTo 10
     End If
     '
     Call CIERRARCH("!SELCUECO")
     Call FRESHECHO("!SELCUECO")
     '
     Call SELECHO("!SELCUECO")
     If VALACT1$("!SELCUECO") = "" Then GoTo 10
     REGCUEC& = Val(VALACT1$("!SELCUECO"))
     If REGCUEC& < 1 Then GoTo 10
     '
     XX0$ = REGLEIDO$("COBRAN", REGCUEC&)
     IMPO# = CVD(Mid$(X$, 15, 8))
     IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     'COMPB$ = String$(18, " "): COMPB$ = "RC-0000-" & MKI$(X$, 3, 4)
     X$ = COMPB$
     Y$ = IMPOS$
     DOCAPLI$ = X$
     '
     If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Abandona\Confirma Desaplicación") <> 2 Then GoTo 10
     '
     ' DESAPLICACION
     ' *************
     '
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("APLICOB")
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
         DODESAP$ = Mid$(X$, 5, 18)
         If DODESAP$ = DOCAPLI$ Then
           Call REPLA(X$, String$(256, 0), 1, 256)
           Call GRAREG("APLIPAG", X$, U&)
         End If
       End If
     Next U&
     Call CIERRARCH("APLIPAG")
     Call BAJALDISCO
     '
     Screen.MousePointer = 1
     If TICUEPRO%((-1) * NC%) = 0 Then
        TTXX$ = "Desaplicacion Completa.\  \Desea Reconstruir la Composición de Deuda\de la Cuenta '" + TRIM$(RASOCLI$((-1) * NC%)) + "' ?"
        If COMALTER%(TTXX$ + "\\Reconstruir\Continuar") = 1 Then
           'Call REAPLICREPRO(NC%)
        End If
     End If
     '
     GoTo 10
     '
End Sub


