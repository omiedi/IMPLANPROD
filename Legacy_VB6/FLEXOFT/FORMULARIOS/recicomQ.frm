VERSION 5.00
Begin VB.Form RECICOMQ 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza - Opciones"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   26
      Top             =   3780
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
      Height          =   855
      Left            =   4095
      Picture         =   "recicomQ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Listados de Recibos Registrados"
      Top             =   4095
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
      Picture         =   "recicomQ.frx":030A
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
      Height          =   855
      Left            =   4095
      Picture         =   "recicomQ.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
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
            Picture         =   "recicomQ.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Acceso Directo a Tabla de Talonarios de Recibo y Cobradores"
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
            Picture         =   "recicomQ.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores"
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
            Picture         =   "recicomQ.frx":09F2
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
            Picture         =   "recicomQ.frx":0B3C
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
            Picture         =   "recicomQ.frx":0C86
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
            Left            =   840
            TabIndex        =   15
            Top             =   80
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
            Picture         =   "recicomQ.frx":0DD0
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
            Picture         =   "recicomQ.frx":143A
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
      Height          =   555
      Left            =   3885
      Picture         =   "recicomQ.frx":1744
      Stretch         =   -1  'True
      Top             =   5120
      Width           =   5250
   End
End
Attribute VB_Name = "RECICOMQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    LU$ = LUDAT$
    If EXISTE%(LU$ + "MECOBRA.DAT") = 1 Then
       If COPYFILE%(LU$ + "MECOBRA.DAT", "C:\FLEXOFT\MECOBRA.WKF") <> 1 Then
          Command1.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COMECOB/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       X$ = REGLEIDO$("!MECOBRA", U&)
       CDX% = CVI(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             J& = J& + 1
363          Call GRAREG("!MECOBRA", X$, J&)
             If TRIM$(Mid$(X$, 3, 12)) = "" Then ATX% = 1
             If TRIM$(Mid$(X$, 45, 12)) = "" Then ATX% = 1
364          If TRIM$(Mid$(X$, 57, 4)) = "" Then ATX% = 1
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MECOBRA", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        Call COMUNI("Datos Incompletos !!!")
        GoTo 360
    End If
    '
    AA% = COPYFILE%("C:\FLEXOFT\MECOBRA.WKF", LU$ + "MECOBRA.DAT")
    Call FINAL("*SOMECO1")
    '
390 Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
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
    'If DERACCE%("REPIRECI", "Repetir Impresion de Recibo Emitido") >= 1 Then
       Call REPIRECI
    'End If
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
    Dim MCOBRA%(1000)
    '
    Call CIERRARCH("COBRADOR")
    Call CIERRARCH("!COBRADOR")
    LU$ = LUDAT$
    If EXISTE%(LU$ + "COBRADOR.DAT") = 1 Then
       If COPYFILE%(LU$ + "COBRADOR.DAT", "C:\FLEXOFT\COBRADOR.WKF") <> 1 Then
          Command6.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COBRADOR/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!COBRADOR")
       X$ = REGLEIDO$("!COBRADOR", U&)
       CDX% = Asc(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 255 Then
             If MCOBRA%(CDX%) = 0 Then
                J& = J& + 1
363             Call GRAREG("!COBRADOR", X$, J&)
                MCOBRA%(CDX%) = 1
             End If
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!COBRADOR", J&)
    Call CIERRARCH("!COBRADOR")
    Call CIERRARCH("COBRADOR")
    '
366 Screen.MousePointer = 1
    '
    AA% = COPYFILE%("C:\FLEXOFT\COBRADOR.WKF", LU$ + "COBRADOR.DAT")
    '
390 Command6.Enabled = True
    '
End Sub


Private Sub Command7_Click()
    Call HELPONLINE("RECISTA")
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


