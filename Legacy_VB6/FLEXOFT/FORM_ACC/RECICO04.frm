VERSION 5.00
Begin VB.Form RECICO04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza - Opciones"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   5580
      Left            =   4515
      ScaleHeight     =   5520
      ScaleWidth      =   1320
      TabIndex        =   16
      Top             =   -210
      Width           =   1380
      Begin VB.CommandButton Command6 
         Caption         =   "D-C"
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
         Picture         =   "RECICO04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
         Top             =   3255
         Width           =   640
      End
      Begin VB.CommandButton Command1 
         Height          =   465
         Left            =   105
         Picture         =   "RECICO04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Re-Imputación Contable"
         Top             =   2730
         Width           =   640
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
         Height          =   690
         Left            =   105
         Picture         =   "RECICO04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1215
         Width           =   640
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
         Height          =   690
         Left            =   105
         Picture         =   "RECICO04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   525
         Width           =   640
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
         Height          =   690
         Left            =   105
         Picture         =   "RECICO04.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Listados de Recibos Registrados"
         Top             =   1995
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   4095
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RECICO04.frx":0BB2
         Top             =   4305
         Width           =   1515
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
      Height          =   4335
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   23
         Top             =   2520
         Width           =   2430
         Begin VB.CommandButton Command10 
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
            Picture         =   "RECICO04.frx":2AD4
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Compensación Cliente-Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compensación Cliente-Proveedor"
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
            TabIndex        =   25
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
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
            Picture         =   "RECICO04.frx":2DDE
            Style           =   1  'Graphical
            TabIndex        =   14
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
            TabIndex        =   15
            Top             =   75
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   10
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
            Picture         =   "RECICO04.frx":2F28
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   7
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
            Left            =   1850
            Picture         =   "RECICO04.frx":3072
            Style           =   1  'Graphical
            TabIndex        =   22
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
            Picture         =   "RECICO04.frx":337C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imputación y Des-Aplicación de Créditos y Anticipos"
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
            Left            =   800
            TabIndex        =   9
            Top             =   80
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
            Picture         =   "RECICO04.frx":34C6
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
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
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
            Picture         =   "RECICO04.frx":37D0
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
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4095
         Y1              =   3300
         Y2              =   3300
      End
   End
End
Attribute VB_Name = "RECICO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%

'\\\OT-06-0409-MH-26/09/06///
Sub Command1_Click()
    '
    Command1.Enabled = False
    '
    If DERACCE%("REIMPUVE", "Re-Imputación de Comprobantes de Ventas") < 2 Then
      GoTo 99
    End If
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    DESDEX = CDbl(CVDAT(DES%))
    HASTAX = CDbl(CVDAT(HAS%))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then GoTo 99
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'RCOB')"
    'OR TIPOMOVIM = 'AJUD'
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechConta >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND FechConta <= " & HASTAX & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            Y$ = MKS$(JJ& + 1) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Tiene\Comprobantes de Venta\que Puedan Re-Imputarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
16  Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    IDOCUM# = XVALO(Mid$(VALACT2$("!SELCUERE"), 19, 14))
    '
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(RASOCLI$(NC%))
    '
    Call BLANARCH("!REIMPUTA")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) >= " & DESDEX & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASTAX & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND IDebeComp > 0 "
    '
    Set CABASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    SUIMPU# = 0
    '
    With CABASEL
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Editar Imputación Contable\del Comprobante Elegido")
        GoTo 10
      End If
      '
      Do While Not .EOF
        CUE% = !CueContaI
        IPA# = !IDEBECOMP - !IHABECOMP
        Y$ = REGBLAN$("!REIMPUTA")
        Call REPLA(Y$, CODCUE$(CUE%), 1, 12)
        Call REPLA(Y$, MKD$(IPA#), 13, 8)
        Call REPLA(Y$, MKI$(CUE%), 29, 2)
        Call REGAPP("!REIMPUTA", Y$)
        SUIMPU# = SUIMPU# + IPA#
        .MoveNext
      Loop
    End With
    Call CIERRARCH("!REIMPUTA")
    '
    If Abs(SUIMPU# - IDOCUM#) >= 0.005 Then
      Call MENSERR(24, "Imposible Editar Imputación Contable.\Comprobante No Balanceado.\  \Debe Anularse y Registrarse Nuevamente.")
      GoTo 10
    End If
    '
30  ATRI$ = "TITUPAN=Re-Imputación Contable: " + ATRI$
    VTB% = VENTABU%("REIMPUTA/" + ATRI$)
    If VTB% < 0 Then GoTo 10
    '
    SUMANE# = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      SUMANE# = SUMANE# + IMPOX#
      '
      If CUENUE% < 1 Then
        Call MENSERR(24, "Numero de Cuenta no Valido\en Registro Numero " + TRIM$(Str$(U&)))
        GoTo 30
      End If
    Next U&
    If Abs(SUMANE# - IDOCUM#) >= 0.005 Then
      Call MENSERR(24, "Comprobante no Balanceado")
      GoTo 20
    End If
    '
    REIMPU% = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      CUEVIE% = CVI(Mid$(Y$, 29, 2))
      '
      With CABASEL
        .MoveFirst
        Do While Not .EOF
32        CUE% = !CueContaI
          If Abs(!IDEBECOMP - !IHABECOMP - IMPOX#) < 0.005 Then
            If CUE% = CUEVIE% Then
              .Edit
              !CueContaI = CUENUE%
                If CUENUE% <> CUEICARCH% Then
                  !Status = 0
                Else
                  ' LO INCORPORA A LA CARTERA PONIENDO STATUS = 1
                  ' SI SE CAMBIO A LA CUENTA CONTABLE DE CARTERA
                  If !Status = 0 Then !Status = 1
                End If
              .Update
              REIMPU% = 1
              GoTo 35
            End If
          End If
        .MoveNext
        Loop
      End With
      '
35  Next U&
    '
    If REIMPU% > 0 Then
        Call COMUNI("Re-Imputación Completada")
      Else
        Call COMUNI("No se Pudo Completar la Re-Imputación")
    End If
    '
    GoTo 16
    '
99  Command1.Enabled = True
    '
End Sub

Sub Command10_Click()
  '\\\OT-06-0112-WAR-30/03/06///
  Command10.Enabled = False
    Call COMPENSACP
  Command10.Enabled = True
  '\\\OT-06-0112-WAR-FIN///
End Sub

Sub Command11_Click()
    Command11.Enabled = False
    OPX% = COMALTER%("Opciones de Consulta de Aplicaciones:\\Por Recibo o N/C\Por Factura o N/D\Listado por Rango de Fechas")
    If OPX% = 1 Then
        Call VAPLIREC
      ElseIf OPX% = 2 Then
        Call VAPLIFAC
      ElseIf OPX% = 3 Then
        Call LIAPLIFE
    End If
    Command11.Enabled = True
End Sub
'\\\OT-06-0409-MH-FIN///

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
       Call REPIRECIBO
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
    TC = TICAMONE(-1) ' fuerza refresco de tabla
    If DERACCE%("IMPUCRED", "Imputacion Créditos/Anticipos Deudores") >= 2 Then
      OPX% = COMALTER%("Opciones de Trabajo y Procesamiento:\\Imputar Créditos\Des-Aplicar Recibos\Cancelar")
      If OPX% = 1 Then
          Hide
          Call FRESHALL
          Call IMPUCREVEN
        ElseIf OPX% = 2 Then
          Call DESAPLIREC
      End If
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

Sub Command6_Click()
   Command6.Enabled = False
   Call OPVARFA04.Command2_Click
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
3  cx% = ULTREG&("DEUDOR1")
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
200 For CX1% = 1 To cx%
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
      DEBE# = IPEN#: Habe# = 0
      If Sgn(IPEN#) < 0 Then DEBE# = 0: Habe# = Abs(IPEN#)
      '
      z$ = REGBLAN$("LISALIN")
      Call REPLA(z$, MKI$(NC%), 1, 2)
      Call REPLA(z$, NROX$, 3, 10)
      Call REPLA(z$, MKI$(FF%), 13, 2)
      Call REPLA(z$, MKD$(IORI#), 15, 8)
      Call REPLA(z$, MKI$(FVI%), 23, 2)
      Call REPLA(z$, MKD$(DEBE#), 25, 8)
      Call REPLA(z$, MKD$(Habe#), 33, 8)
      Call REPLA(z$, MKD$(0), 41, 8)
      '
      Call REGAPP("LISALIN", z$)
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
   Call LISTADORECIBOS_SUBDIARIORETENCIONES_ATRASOCOBRANZA(0)
   Command8.Enabled = True
   '
End Sub
Sub LISTADORECIBOS_SUBDIARIORETENCIONES_ATRASOCOBRANZA(OPXX%)
   If OPXX% < 1 Then
      OPXX% = ALTERNA%("Listados de Recibos Registrados\Subdiario de Retenciones de Ingresos Brutos\Atraso de Cobranza")
   End If
   '
   If OPXX% = 1 Then
       GoTo 3
   ElseIf OPXX% = 2 Then
       Call LISRETIIBB
       GoTo 99
   ElseIf OPXX% = 3 Then
       Call ATRACOBRA
       GoTo 99
   Else
       GoTo 99
   End If
     
3  Dim RERECI&(10000)
   '
'10 FEX = HOY(HO$)
'   HAS% = FEX
'   HASS$ = FECHATEX$(FEX)
'   DESS$ = "01" + Mid$(HASS$, 3)
'   DES% = FECHANUM(DESS$)
'   VDEF$ = MKI$(DES%) + MKI$(HAS%)
'   FEX = DES%: PERIO$ = FECHATEX$(FEX)
'   FEX = HAS%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
'
'   '
'   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
'   If OBX$ = "" Then
'      GoTo 99
'   End If
'   DES% = CVI(Left$(OBX$, 2))
'   HAS% = CVI(Mid$(OBX$, 3, 2))
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then GoTo 99

   '
   OPXX% = COMALTER%("Opciones de Ordenamiento:\\Por Fecha\Por Comprobante")
   If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
20 NUSUC% = 0
   If ULTREG&("COBRADOR") > 0 Then
      Call SELECHO("COBRADOR")
      TT2$ = VALACT2$("COBRADOR")
      'If TRIM$(TT2$) = "" Then GoTo 10
      NUSUC% = XVALO(TRIM$(VALACT1$("COBRADOR")))
      'If NUSUC% < 0 Then GoTo 10
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = Int(CDbl(CVDAT(DES%)))
   HASA = Int(CDbl(CVDAT(HAS%)))
   'nuevo
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (INT(CDBL(FechConta)) >= " & DESA & " "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASA & " "
   If NUSUC% > 0 Then
     SQLX$ = SQLX$ + "AND NumeTalo = " & NUSUC% & ") "
    Else
     SQLX$ = SQLX$ + ")"
   End If
   SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
   SQLX$ = SQLX$ + "Order by FechEmisi ASC "
   '\\\OT-06-0162-WAR-04/05/06///
   SQLX$ = SQLX$ + ", NumeComp ASC "
   '\\\OT-06-0162-WAR-FIN///
   Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   'ESCRITURA
   With AVANCETMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
       GoTo 99
     End If
     JJ& = 0
     REBLAY$ = REGBLAN$("AVAFACB")
     Do While Not .EOF
       
       NC% = !NuClien
       CLIEX$ = ""
       If NC% > 0 Then
          CLIEX$ = RASOCLI$(NC%)
          NRO& = !NUMECOMP
          NETO# = !ValAbsComp 'Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
            If !OPCIMOVIM = "NC" Then NETO# = (-1) * NETO#
          NUSUCUVE% = !CodiSucu
          NNX$ = !CODICOM_LAR
          '
          MONEM% = 0: TICAM = 1
          On Error Resume Next
          MONEM% = !Mone_Emis
          On Error GoTo 0
          '
          TICAM = 1
          On Error Resume Next
          If IsNull(!Tipo_Cam) = False Then
              TICAM = !Tipo_Cam
          End If
          On Error GoTo 0
          '
          Y$ = REBLAY$
          Call REPLA(Y$, MKI$(CVINT(!FechConta)), 1, 2)
          Call REPLA(Y$, CLIEX$, 3, 28)
          Call REPLA(Y$, NNX$, 31, 18)
          Call REPLA(Y$, MKD$(NETO#), 49, 8)
          '
          If MONEM% > 1 Then
             If TICAM > 0.1 Then
                Call REPLA(Y$, SIMBOLPES$(MONEM%), 65, 4)
                Call REPLA(Y$, MKD$(NETO# / TICAM), 69, 8)
             End If
          End If
          '
          JJ& = JJ& + 1
          Call GRAREG("AVAFACB", Y$, JJ&)
       End If
49      .MoveNext
     Loop
   End With
   '
   Call SETULTREG("AVAFACB", JJ&)
   
   Call CIERRARCH("*.*")
   If OPXX% = 1 Then Call FILESORT("AVAFACB", "", 1, 1, "ASC")
   '
   If OPXX% = 2 Then Call FILESORT("AVAFACB", "", 3, 3, "ASC")
   Call HEADERS("AVAFACB", "")
   Call HEADERS("AVAFACB", "Periodo: " + PERIO$ + " - Todos Los Vendedores")
   If TT2$ <> "" Then Call HEADERS("AVAFACB", "Talonario: " + TT2$)
   '
   If JJ& < 1 Then
     Call MENSERR(24, "Sin Movimientos\en el Rango de Fechas Elegido")
     GoTo 10
   End If
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Call CONECRUN("*LIAVAFB", "Secuencia de Recibos Registrados/WAIT")
   If ULTREG&("COBRADOR") > 1 Then
     If NUSUC% > 0 Then GoTo 20
   End If
   '
99 Screen.MousePointer = 1
End Sub
Sub ANURECIBO()
    '
    If DERACCE%("ANURECO", "Anulación de Recibos de Cobranza") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Cliente
    '     ****************************
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'RCOB' OR (TipoMovim = 'AJUD' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            JJ& = JJ& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(JJ&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Comprobantes de Cobranza \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If TRIM$(VALACT2$("!SELCUERE")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DOCULB$ = "AS" + TRIM(Mid(DOCULA$, 3, 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    '\\\OT-05-0736-WAR-11/11/05/// NO ANULA SI CHEQUE ENTREGADO O DEPOSITADO
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    '
    Set CABANSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    With CABANSEL
      SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NC% & " "
      .FindFirst SQLX$
15    If .NoMatch = False Then
        If UCase$(!OPCIMOVIM) <> "CH" Then GoTo 17 ' NO ES MOVIMIENTO DE CHEQUE
        If !Status > 1 Then GoTo 16               ' SALIDO DE CARTERA
        If IsNull(!FECHENTRE) = True Then GoTo 17  ' NO TIENE FECHA DE ENTREGA
        If CVINT(!FECHENTRE) < 33 Then GoTo 17     ' LA FECHA DE ENTREGA ES MENOR A 33
        '
16      Call MENSERR(24, "Anulación Imposible.\ Valores Ingresados con este Recibo\Figuran Depositados o Entregados en Pago.")
        GoTo 999
        '
17      .FindNext SQLX$
        GoTo 15
      End If
    End With
    '\\\OT-05-0736-WAR-FIN///
    '
18  RESERNUM% = COMALTER%("Desea Dejar Reservado\este Número de Comprobante ?\\Reservar Número\Re-Utilizarlo")
    If RESERNUM% < 1 Or RESERNUM% > 2 Then GoTo 18
    '
    ' BORRA LOS CHEQUES ASOCIALDOS AL RECIBO
    SQLX$ = "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "')"
    SQLX$ = SQLX$ + " AND NuClien = " & NC% & " "
    If RESERNUM% = 1 Then
       SQLX$ = SQLX$ + " AND TRIM(TipoMovim) <> 'RCOB'"
    End If
    On Error Resume Next
    CueCorri.Execute "DELETE * FROM CajaBanc WHERE " & SQLX$
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Borrar los Cheques - Consulte.")
    End If
    On Error GoTo 0
    '
    ' PONE EN CERO EL RECIBO
    SQLX$ = "SELECT * FROM CajaBanc "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "')"
    SQLX$ = SQLX$ + " AND NuClien = " & NC% & " "
    Set MIRS = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    On Error Resume Next
    With MIRS
      .MoveFirst
19    If Err < 1 Then
        .Edit
        !ValAbsComp = 0
        !IDEBECOMP = 0
        !IHABECOMP = 0
        !BaseImpo = 0
        !IdenBaja = IDACTU$
        !FeHoReal = Now
        .Update
        .MoveNext
        GoTo 19
      End If
    End With
    On Error GoTo 0
    '
    ' BORRA LAS APLICACIONES
    SQLX$ = "(CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "OR CoDocOr_Lar = '" & DOCULB$ & "' OR CoDocAp_Lar = '" & DOCULB$ & "') "
    SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
    On Error Resume Next
    CueCorri.Execute "DELETE * FROM ApliCobra WHERE " & SQLX$
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Borrar las Aplicaciones - Consulte.")
    End If
    On Error GoTo 0
    '
'    SQLX$ = "SELECT * FROM APLICOBRA "
'    SQLX$ = SQLX$ + "WHERE (CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
'    SQLX$ = SQLX$ + "OR CoDocOr_Lar = '" & DOCULB$ & "' OR CoDocAp_Lar = '" & DOCULB$ & "') "
'    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
'    Set APLICSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
'
'    wrkPredeterminado.BeginTrans
'    With CABANSEL
'      .FindFirst "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NC% & " "
'21    If .NoMatch = False Then
'        If RESERNUM% <> 1 Or TRIM$(!TipoMovim) <> "RCOB" Then
'             .Delete
'           Else
'             .Edit
'             !ValAbsComp = 0
'             !IDEBECOMP = 0
'             !IHABECOMP = 0
'             !BASEIMPO = 0
'             !IdenBaja = IDACTU$
'             !FeHoReal = Now
'             .Update
'        End If
'        .FindNext "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NC% & " "
'        GoTo 21
'      End If
'    End With
'    '
'    With APLICSEL
'31    .FindFirst "NuClien = " & NC% & " "
'      If .NoMatch = False Then
'         .Delete
'         GoTo 31
'      End If
'    End With
'
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If (TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Or TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULB$)) Then
        DOCABE$ = TRIM$(Mid$(XXC$, 13, 18))
        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
      End If
    Next KUL&
'    wrkPredeterminado.CommitTrans
    '
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    RECUEC04.Text1.TEXT = ""
    RECUEC04.Text1.TEXT = TRIM$(Str$(NC%))
    On Error Resume Next
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    RECUEC04.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Sub AJUSALCLI_RODRIGO()
    '
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "BIMONETA" Or MODOMONE$ = "MULTIMON" Then
      OPXX% = ALTERNA%("Ajuste Cuenta Corriente\Saldo Factura Moneda Extranjera")
      If OPXX% = 1 Then
          GoTo 3
        ElseIf OPXX% = 2 Then
          Call AJUSALDOL
          Exit Sub
        Else
          Exit Sub
      End If
    End If
    '
3   Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%) + MKI$(1) 'TOMA POR DEFECTO LA MONEDA PESOS
    CLSS$ = OBJPLA$("SELCLIAJU1", VDEF$) 'ABRE EL NUEVO OBJETO CREADO 'SELCLIAJU1' QUE INCLUYE
    '
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$(NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    '
    MONE% = CVI(Mid$(CLSS$, 5, 2))
    '
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
'2.- Determina Saldo por arrastre de Facturas y Debitos
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        IFAC# = !IDEBECOMP - !IHABECOMP
        FF% = CVINT(!FechConta)
        SALFINAL# = SALFINAL# + IFAC#
        If FF% <= FAJUI% Then
          SARRASFE# = SARRASFE# + IFAC#
        End If
      .MoveNext
      Loop
    End With
    '
'3.- Determina saldo por composicion de deuda
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "where NuClien = " & NC% & " "
    '
    Set CODEUTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CODEUTEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        SALDEUTO# = SALDEUTO# + !SaldDebe - !SaldAcre
      .MoveNext
      Loop
    End With
    '
100 SALAJU# = SARRASFE#
    '
    '20/03/07 (OT-07-0120)'MUESTRA Y LLENA EL OBJETO 'CARAJUCLI1'
    If MONE% > 1 Then
       MONDES$ = TRIM$(ECOARCH$("TAMONED", MKI$(MONE%)))
       COTIZ# = TICAMONE(MONE%)
       VDEF1$ = Space(71)
       Call REPLA(VDEF1$, MKI$(NC%), 1, 2)
       Call REPLA(VDEF1$, MKI$(FAJUI%), 3, 2)
       Call REPLA(VDEF1$, MKI$(0), 21, 2)
       Call REPLA(VDEF1$, MKI$(0), 23, 2)
       Call REPLA(VDEF1$, MKI$(MONE%), 25, 2)
       Call REPLA(VDEF1$, TRIM$(MONDES$), 27, 30)
       Call REPLA(VDEF1$, MKD$(COTIZ#), 65, 8)
101    CLSS$ = OBJPLA$("CARAJUCLI1", VDEF1$)
       If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
          SALAJU# = CVD(Mid$(CLSS$, 13, 8))
          ICONI% = CVI(Mid$(CLSS$, 279, 2))
          If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
             Call MENSERR(24, "Falta Imputación Contable Ajuste")
             VDEF1$ = CLSS$
             GoTo 102
          End If
          If ICONI% = CUECOCLI%(NC%) Then
             Call MENSERR(24, "Cuenta Contable no Válida")
             VDEF1$ = CLSS$
             GoTo 101
          End If
    '20/03/07 (FIN)
    '
    Else
       VDEF$ = Mid$(CLSS$, 1, 4) + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0)
102    CLSS$ = OBJPLA$("CARAJUCLI", VDEF$)
       If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
          SALAJU# = CVD(Mid$(CLSS$, 13, 8))
          ICONI% = CVI(Mid$(CLSS$, 279, 2))
          If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
             Call MENSERR(24, "Falta Imputación Contable Ajuste")
             VDEF$ = CLSS$
             GoTo 102
          End If
          If ICONI% = CUECOCLI%(NC%) Then
             Call MENSERR(24, "Cuenta Contable no Válida")
             VDEF$ = CLSS$
             GoTo 102
          End If
    End If
    '
    DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
    DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
    '
    FVENI% = CVI(Mid$(CLSS$, 21, 2))
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    '
    ORMA% = (-1)
    Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECITEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
      End If
    End With
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
20  NC1% = NC%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = NC%
    NumeProv% = 0
    FechEmi = CVDAT(FAJUI%)
    '
    Referen$ = RASOCLI$(NC%)
    FeContab = CVDAT(FAJUI%)
    FeAcredi = CVDAT(FAJUI%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = RASOCLI$(NC%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC%))
    TiMovi = "AJUD"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + RASOCLI$(NC%)
    ImpoDeb# = DIFESAL#
    '\\\OT-05-0493-WAR-11/07/05///
    HAYAJU% = 0: If Abs(DIFESAL#) > 0.005 Then HAYAJU% = 1
    '\\\OT-05-0493-WAR-FIN///
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALD"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "AJU-RO"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + RASOCLI$(NC%)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FVENI%)
    ImVenc# = DIFEDEU#
    Call GRADEUPEN
    '
    CajaBanc.Close
    '
    
    If HAYAJU% > 0 Then
      '
      MOTIVO$ = MOTIVO$ + Mid$(CLSS$, 23, 279)
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
         NUMEDOC$ = NRECIB$ 'TRIM$(Str$(NUAJU&)) '
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
    End If
    '
    Screen.MousePointer = 1
    Call COMUNI("Ajuste Completo")
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'
Sub AJUSALCLI()
    '
    If MODOMONEI% = 2 Then
      OPXX% = ALTERNA%("Ajuste Cuenta Corriente\Saldo Factura Moneda Extranjera")
      If OPXX% = 1 Then
          GoTo 3
        ElseIf OPXX% = 2 Then
          Call AJUSALDOL
          Exit Sub
        Else
          Exit Sub
      End If
    End If
    '
3   Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   Screen.MousePointer = 1
    VDEF$ = MKI$(NC%) + MKI$(FAJUI%) 'TOMA POR DEFECTO LA MONEDA PESOS
    CLSS$ = OBJPLA$("SELCLIAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$(NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
'2.- Determina Saldo por arrastre de Facturas y Debitos
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        IFAC# = ROUND#(!IDEBECOMP - !IHABECOMP)
        FF% = CVINT(!FechConta)
        SALFINAL# = SALFINAL# + IFAC#
        If FF% <= FAJUI% Then
          SARRASFE# = SARRASFE# + IFAC#
        End If
      .MoveNext
      Loop
    End With
    '
'3.- Determina saldo por composicion de deuda
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "where NuClien = " & NC% & " "
    '
    Set CODEUTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CODEUTEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        SALDEUTO# = SALDEUTO# + ROUND#(!SaldDebe - !SaldAcre)
      .MoveNext
      Loop
    End With
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0) + Space$(256) + MKI$(0)
    OBJE$ = "CARAJUCLI"
    If MODOMONEI% = 2 Then
       OBJE$ = "CARAJUCLI-DOL"
       VDEF$ = VDEF$ + MKI$(1) + MKS$(1)
    End If
102 CLSS$ = OBJPLA$(OBJE$, VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 13, 8))
    ICONI% = CVI(Mid$(CLSS$, 279, 2))
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
       Call MENSERR(24, "Falta Imputación Contable Ajuste")
       VDEF$ = CLSS$
       GoTo 102
    End If
    If ICONI% = CUECOCLI%(NC%) Then
       Call MENSERR(24, "Cuenta Contable no Válida")
       VDEF$ = CLSS$
       GoTo 102
    End If
    '
    DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
    DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
    '
    If Abs(DIFESAL#) < 0.005 Then
      If Abs(DIFEDEU#) < 0.005 Then
        Call MENSERR(24, "No hay Ajuste")
        GoTo 5
      End If
    End If
    '
    MONAJU% = 1 ' PESOS
    COTIMONEM# = 1
    If MODOMONEI% = 2 Then
      MONAJU% = CVI(Mid$(CLSS$, 281, 2))
      If MONAJU% < 1 Or MONAJU% > 6 Or ECOARCH$("TAMONED", MKI$(MONAJU%)) = "" Then
         Call MENSERR(24, "Moneda no Válida")
         VDEF$ = CLSS$
         GoTo 102
      End If
      '
      COTIMONEM# = CVS(Mid$(CLSS$, 283, 4))
      If COTIMONEM# < 0.01 Then
         Call MENSERR(24, "Falta Cotización")
         VDEF$ = CLSS$
         GoTo 102
      End If
      If MONAJU% = 1 Then
         If Abs(COTIMONEM# - 1) > 0.00001 Then
            Call MENSERR(24, "Cotización no Válida")
            VDEF$ = CLSS$
            GoTo 102
         End If
      End If
    End If
    '
    FVENI% = CVI(Mid$(CLSS$, 21, 2))
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    '
    ORMA% = (-1)
    Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECITEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
      End If
    End With
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    MONEMIA% = MONEMI%
    '
20  NC1% = NC%
    Call APERBASDAT("CABAN")
    '
    MONEMI% = MONAJU%
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = NC%
    NumeProv% = 0
    FechEmi = CVDAT(FAJUI%)
    '
    Referen$ = RASOCLI$(NC%)
    FeContab = CVDAT(FAJUI%)
    FeAcredi = CVDAT(FAJUI%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = RASOCLI$(NC%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC%))
    TiMovi = "AJUD"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + RASOCLI$(NC%)
    ImpoDeb# = DIFESAL#
    ValoDola# = ROUND#(DIFESAL# / COTIMONEM#)
    HAYAJU% = 0: If Abs(DIFESAL#) > 0.005 Then HAYAJU% = 1
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALD"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "AJU-CL"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + RASOCLI$(NC%)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FVENI%)
    ImVenc# = DIFEDEU#
    Call GRADEUPEN
    '
    CajaBanc.Close
    '
    If MONEMI% = 2 Then
                      Call REPLA(YCB$, MKI$(FAJUI%), 1, 2)
                      Call REPLA(YCB$, MKI$(NC%), 3, 2)
                      Call REPLA(YCB$, NRECIB$, 13, 18)
                      Call REPLA(YCB$, MKS$(NumeCom), 7, 4)
                      Call REPLA(YCB$, MKI$(MONEMI%), 31, 2)
                      Call REPLA(YCB$, MKI$(NC%), 35, 2)
                      Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                      Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                      Call REPLA(YCB$, MKS$(NumeCom), 39, 4)
                      Call REPLA(YCB$, MKI$(0), 43, 2)
                      Call REPLA(YCB$, NRECIB$, 45, 18)
                      Call REPLA(YCB$, MKI$(MONEMI%), 63, 2) ' RECIBO SIEMPRE EN PESOS
                      Call REPLA(YCB$, MKD$(ROUND#(DIFESAL#)), 65, 8)
                      Call REPLA(YCB$, MKD$(0), 73, 8)
                      Call REPLA(YCB$, MKD$(0), 81, 8)
                      Call REPLA(YCB$, MKD$(COTIMONEM#), 89, 8)
                      Call REPLA(YCB$, MKD$(ROUND#(DIFESAL# / COTIMONEM#)), 97, 8)
                      Call REPLA(YCB$, MKD$(0), 105, 8)
                      Call REPLA(YCB$, MKD$(0), 113, 8)
                      'Call REPLA(YCB$, String$(32, 0), 89, 32)
                      Call REGAPP("CARBIMON", YCB$)
    End If
    
    If HAYAJU% > 0 Then
      '
      MOTIVO$ = MOTIVO$ + Mid$(CLSS$, 23, 279)
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
         NUMEDOC$ = NRECIB$ 'TRIM$(Str$(NUAJU&)) '
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
    End If
    '
    MONEMI% = MONEMIA%
    Screen.MousePointer = 1
    Call COMUNI("Ajuste Completo")
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'
Sub AJUSALDOL()
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Dim DOCUCA$(), SALPES#(), SALDOL#()
    CADOC% = 0
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For JKLY& = 1 To Len(RFF$) Step 4
      RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
      YCB$ = REGLEIDO$("CARBIMON", RECABI&)
      If TRIM$(Mid$(YCB$, 13, 18)) <> "" Then
        If TRIM$(Mid$(YCB$, 13, 18)) = TRIM$(Mid$(YCB$, 45, 18)) Then
          CADOC% = CADOC% + 1
          ReDim Preserve DOCUCA$(CADOC%), SALPES#(CADOC%), SALDOL#(CADOC%)
          DOCUCA$(CADOC%) = TRIM$(Mid$(YCB$, 13, 18))
          SALPES#(CADOC%) = 0
          SALDOL#(CADOC%) = 0
        End If
      End If
    Next JKLY&
    '
    For JKLY& = 1 To Len(RFF$) Step 4
      RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABI&)
      DOCABE$ = TRIM$(Mid$(XXC$, 13, 18))
      For KUI% = 1 To CADOC%
        If DOCUCA$(KUI%) = DOCABE$ Then
          SALPES#(KUI%) = SALPES#(KUI%) + (Int(100 * (CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))) + 0.5)) / 100
          SALDOL#(KUI%) = SALDOL#(KUI%) + (Int(100 * (CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))) + 0.5)) / 100
          GoTo 19
        End If
      Next KUI%
19  Next JKLY&
    '
    JJ& = 0
    For KUI% = 1 To CADOC%
      If Abs(SALPES#(KUI%)) < 0.005 Then
        If Abs(SALDOL#(KUI%)) >= 0.005 Then
          IMPOS$ = NUMSTRI$(SALDOL#(KUI%), 14, 2, 1)
          NPX& = XVALO(Mid$(DOCUCA$(KUI%), 11, 8))
          Y$ = MKS$(NPX&) + AJUSTI$(DOCUCA$(KUI%), 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      End If
    Next KUI%
    Call SETULTREG("!SELCUERE", JJ&)
    Call FRESHECHO("!SELCUERE")
    '
    If JJ& < 1 Then
      Call COMUNI("La Cuenta Elegida no Presenta\Comprobantes en Moneda Extranjera\que Puedan Ajustarse.")
      GoTo 10
    End If
    '
    Call SELECHO("!SELCUERE")
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If XVALO(REGCUEC&) < 1 Then
        GoTo 10
    End If
    '
    DOCUCABE$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DIFEX# = XVALO(Mid$(VALACT2$("!SELCUERE"), 19, 14))
    If COMALTER%("Confirme el Ajuste a CERO del Saldo\en Moneda Extranjera del Documento '" + DOCUCABE$ + "'.\\Confirmar\Cancelar") = 1 Then
      RECILAR$ = "AJ-X-0000-00000000"
      RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
      For JKLY& = 1 To Len(RFF$) Step 4
        RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
        YCB$ = REGLEIDO$("CARBIMON", RECABI&)
        If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
            MONEMORI% = CVI(Mid$(YCB$, 31, 2))
            TICAMBX# = TICAMONE(MONEMORI%)
            FEDOAPLI% = HOY(HOS$)
            NROO& = XVALO(Mid$(RECILAR$, 11))
          Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
          Call REPLA(YCB$, MKI$(NC%), 35, 2)
          Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
          Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
          Call REPLA(YCB$, MKS$(NROO&), 39, 4)
          Call REPLA(YCB$, MKI$(SUC%), 43, 2)
          Call REPLA(YCB$, RECILAR$, 45, 18)
          Call REPLA(YCB$, MKI$(MONEMORI%), 63, 2) ' RECIBO SIEMPRE EN PESOS
          Call REPLA(YCB$, MKD$(0), 65, 8)
          Call REPLA(YCB$, MKD$(0), 73, 8)
          Call REPLA(YCB$, MKD$(0), 81, 8)
          Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
          Call REPLA(YCB$, String$(16, 0), 97, 16)
          If DIFEX# < 0 Then
              Call REPLA(YCB$, MKD$(Abs(DIFEX#)), 97, 8)
            Else
              Call REPLA(YCB$, MKD$(DIFEX#), 105, 8)
          End If
          Call REPLA(YCB$, MKD$(0), 113, 8)
          '
          Call REGAPP("CARBIMON", YCB$)
          Call CIERRARCH("CARBIMON")
          '
          Call RECUSALDOL(DOCUCABE$, NC%)
          '
          Call COMUNI("Ajuste Procesado")
          GoTo 10
        End If
      Next JKLY&
      '
      Call MENSERR(24, "No fue Posible Procesar el Ajuste")
    End If
    '
    GoTo 10
End Sub
'
Sub IMPUCREVEN()
    '
10  If APLIAUTO% = 1 Then
       If NC% > 0 Then
          GoTo 12
       End If
    End If
    On Error GoTo 0
    '
    Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
12  FIMPUC04.Label1 = Str$(NC%)
    FIMPUC04.Label2 = RASOCLI$(NC%)
    FIMPUC04.Label3 = DOMICLI$(NC%)
    FIMPUC04.Label6 = CPOSCLI$(NC%)
    FIMPUC04.Label12 = LOCACLI$(NC%)
    '
    FIMPUC04.List2.Clear
    Call BLANARCH("!SELDOCAP")
    Call FRESHECHO("!SELDOCAP")
    FIMPUC04.Refresh
    '
' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
          IMPO# = !IBruComp
          'SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = !SaldDebe - !SaldAcre
          SALDOL# = !SaldDebe_Dola - !SaldAcre_Dola
          MEMI% = !Mone_Emis
          'SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = !CODICOM_LAR
          FEMI = CVINT(!FechEmisi)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
          '
          If SALDO# >= 0.005 Or SALDOL# > 0 Then
               '
               SUMAFAC# = SUMAFAC# + IMPO#
               IMPO# = Abs(IMPO#)
               SUMADEU# = SUMADEU# + SALDO#
               '
               z$ = Space$(128)
               Call REPLA(z$, NROX$, 1, 18)
               Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
               Call REPLA(z$, IMPTEX$, 32, 13)
               Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
               Call REPLA(z$, SALTEX$, 57, 13)
               '        Call REPLA(Z$, REGTEX$, 119, 8)
               '        Call REPLA(Z$, IRETEX$, 127, 2)
               If MODOMONEI% = 2 Then GoTo 20
               '
               ' ESTO ES PARA PROCESO BIMONETARIO
               MONEMORI% = 0: SALDOLAX# = 0
               RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
               DOCUCABE$ = TRIM$(NROX$)
               For JKLY& = 1 To Len(RFF$) Step 4
                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                   If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                   End If
                   SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
                   Call REPLA(YCB$, MKD$(0), 113, 8)
                 End If
               Next JKLY&
               If MONEMORI% > 1 Then
                 SALDO1# = SALDOLAX# * TICAMONE(MONEMORI%)
                 If SALDO1# < 0 Then SALDO1# = 0
                 If SALDO1# > SALDO# Then SALDO1 = SALDO#
                 SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 12, 2, 2)
                 Call REPLA(z$, SALTEX$, 57, 13)
               End If
               ' FIN ESTO ES PARA PROCESO BIMONETARIO
               '
20             FIMPUC04.List2.AddItem z$
               '
             ElseIf SALDO# <= (-0.005) Or SALDOL# <= (-0.005) Then
               '
               z$ = Space$(60) + MKS$(0)
               Call REPLA(z$, NROX$, 1, 18)
               Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
               Call REPLA(z$, IMPTEX$, 27, 14)
               Call REPLA(z$, SALTEX$, 41, 14)
               '     Call REPLA(Z$, MKS$(REG&), 61, 4)
               '     Call REPLA(Z$, MKI$(KKI%), 59, 2)
                    '
               Call REGAPP("!SELDOCAP", z$)
               KQ& = KQ& + 1
               '
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    If ULTREG&("!SELDOCAP") < 1 Then
      Call MENSERR(24, "No Hay Créditos Pendientes de Aplicación\en la Cuenta Elegida.")
      APLIAUTO% = 0
      GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELDOCAP")
    If Len(SPDX$) <= 4 Then
       APLIAUTO% = 0
       GoTo 10
    End If
    '
    With FIMPUC04
      .Label27 = "Saldo no Imputado:"
      .Label16 = ""
      .Label16.Visible = False
      .Label31.Visible = False
      .Label30.Visible = False
      .Label28.Visible = False
      '
      NROXEL$ = TRIM$(Left$(SPDX$, 18))
      .Label25 = NROXEL$
      .Label24 = Mid$(SPDX$, 19, 8)
      .Label23 = Mid$(SPDX$, 27, 14)
      .Label22 = Mid$(SPDX$, 41, 14)
      .Label14 = Mid$(SPDX$, 41, 14)
      .Label18 = ""
      .Label20 = .Label14
      .Label10 = ""
      If MODOMONEI% = 2 Then
         NCX1% = XVALO(.Label1)
         FEMIX = FECHANUM(.Label24)
         MEMI% = MONEMISOR%(NROXEL$, FEMIX, NCX1%)
         TICAM# = TICAMISOR(NROXEL$, FEMIX, NCX1%)
         .Label23 = FORMATNUM$(ROUND#((XVALO(.Label23)) / TICAM#), "F12.2")
         .Label22 = FORMATNUM$(ROUND#((XVALO(.Label22)) / TICAM#), "F12.2")
         .Label14 = FORMATNUM$(ROUND#((XVALO(.Label14)) / TICAM#), "F12.2")
         .Label27 = "Saldo N/I:"
         .Label16 = FORMATNUM$(TICAM#, "F9.4")
         .Label28 = SIMBOLPES$(MEMI%)
         .Label32 = TRIM$(Str$(MEMI%))
         .Label16.Visible = True
         .Label28.Visible = True
         .Label30.Visible = True
         .Label31.Visible = True
         '
         If MEMI% > 1 Then
            ' VOLVER A CARGAR EL LIST2
            .List2.Clear
            With FacPenTemp
               On Error Resume Next
               .MoveLast
               If Err < 1 Then
                  On Error GoTo 0
                  .MoveFirst
                  Do While Not .EOF
                     TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
                     IMPO# = !IBruComp
                     SALDO# = !SaldDebe_Dola - !SaldAcre_Dola
                     MONEDOCU% = !Mone_Emis
                     If MEMI% > 1 Then
                       If MEMI% <> MONEDOCU% Then GoTo 95
                     End If
                     NROX$ = !CODICOM_LAR
                     FEMI = CVINT(!FechEmisi)
                     FEVE = CVINT(!FEVENCIM)
                     IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
                     SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
                     '
                     If SALDO# >= 0.005 Then
                        z$ = Space$(128)
                        Call REPLA(z$, NROX$, 1, 18)
                        Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
                        Call REPLA(z$, IMPTEX$, 32, 13)
                        Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
                        Call REPLA(z$, SALTEX$, 57, 13)
                        FIMPUC04.List2.AddItem z$
                     End If
95                .MoveNext
                  Loop
               End If
            End With
            ' FIN DE VOLVER A CARGAR EL LIST2
         End If
      End If
    End With
    '
    FIMPUC04.Show 1
    GoTo 10
    '
999 End Sub

Sub DESAPLIREC()
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Call BLANARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT DISTINCT CoDocAp_Lar, FechEmisi FROM ApliCobra "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    On Error Resume Next
    With FacPenTemp
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
        TIX$ = !CoDocAp_Lar
        SUAPLI# = 0
        SQLX$ = "NuClien = " & NC% & " AND CoDocAp_Lar = '" & TIX$ & "'"
          '
          ApliCobra.FindFirst SQLX$
20        If ApliCobra.NoMatch = False Then
            SUAPLI# = SUAPLI# + ApliCobra!ImpoApli
            ApliCobra.FindNext SQLX$
            GoTo 20
          End If
          '
          IMPOS$ = NUMSTRI$(SUAPLI#, 14, 2, 1)
          NPX& = XVALO(Mid$(!CoDocAp_Lar, 11, 8))
          Y$ = MKS$(NPX&) + AJUSTI$(!CoDocAp_Lar, 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
          '
        .MoveNext
        Loop
      End If
    End With
    '
    If ULTREG&("!SELCUERE") < 1 Then
      Call MENSERR(24, "No Hay Créditos Aplicados\en la Cuenta Elegida.")
      GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELCUERE")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    TIX$ = Mid$(SPDX$, 5, 18)
    SQLX$ = "NuClien = " & NC% & " AND CoDocAp_Lar = '" & TIX$ & "'"
    '
    With ApliCobra
30    .FindFirst SQLX$
      If .NoMatch = False Then
        .Delete
        GoTo 30
      End If
    End With
    '
    DOCULA$ = TIX$
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
      End If
    Next KUL&
    '
    Call COMUNI("Des-Aplicación de " + TRIM$(TIX$) + " Completa.\Debe Ahora Reconstruir la Composición de Deuda del Cliente.")
    '
    RECUEC04.Text1.TEXT = TRIM$(Str$(NC%))
    Hide
    Call FRESHALL
    PROLLAMA$ = Name
    Call RECUEC04.Command4_Click
    RECUEC04.Show 1
    Call FRESHALL
    '
    GoTo 10
    '
999 End Sub
'
Sub REPIRECIBO()
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FENUE$ = CONTROL$("", "FENURECI")
    FLIXX% = FECHANUM(FENUE$)
    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE TipoMovim= 'RCOB' "
    SQLX$ = SQLX$ + " AND NuClien= " & NC% & " "
    Set COBRANTMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With COBRANTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      Do While Not .EOF
        NROLL& = !NUMECOMP
        IMPO# = !ValAbsComp
        FC% = CVINT(!FechEmisi)
        '
        If NROLL& > 0 Then
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          NNXX# = CDbl(NROLL&)
          DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
          Call REPLA(Y$, MKS$(!NUMECOMP), 1, 4)
          Call REPLA(Y$, DOCU$, 5, 18)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
        .MoveNext
      Loop
    End With
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
      Call MENSERR(24, "Sin Recibos Posteriores al " + FENUE$)
      GoTo 10
    End If
    '
350 Call SELECHO("!SELCUERE")
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If XVALO(REGCUEC&) < 1 Then
        GoTo 999
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE NumeComp = " & REGCUEC & " "
    SQLX$ = SQLX$ + " AND TipoMovim= 'RCOB'"
    Set COBRANTMPB = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With COBRANTMPB
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Imposible Abrir Documento Desde esta Aplicación")
          GoTo 10
        End If
        On Error GoTo 0
        NCLIE% = !NuClien
        NROLL& = !NUMECOMP
    End With
    '
    TIDOCUM$ = "Recibo Nro."
    NUDOCU$ = " " + TRIM$(Str$(NROLL&))
    RACLI$ = TRIM$(Left$(RASOCLI$(NCLIE%), 8))
    '
    Screen.MousePointer = 11
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          If InStr(XX$, RACLI$) > 16 Then
            DOCUNU& = CVS(Left$(XX$, 4))
            GoTo 5
          End If
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
    GoTo 10
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 10
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
     EDITFORM.Picture1(KKI%).Cls
     EDITFORM.Picture1(KKI%).Top = 0
     EDITFORM.Picture1(KKI%).Refresh
     EDITFORM.Picture1(KKI%).Height = 7170
     EDITFORM.Picture1(KKI%).Width = 11700
     EDITFORM.Picture1(KKI%).DrawWidth = 3
     EDITFORM.Picture1(KKI%).DrawStyle = 0
     EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
96  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 96
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Visible = False
      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    'EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 350
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    '
    Screen.MousePointer = 1
    '







    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
'10  Screen.MousePointer = 1
'    Call SELECHO("DEUDOR1")
'    NC% = XVALO(VALACT1$("DEUDOR1"))
'    If NC% < 1 Then
'        GoTo 999
'    End If
'    '
'    Screen.MousePointer = 11
'    FENUE$ = CONTROL$("", "FENURECI")
'    FLIXX% = FECHANUM(FENUE$)
'    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
'    Call BLOQARCH("!SELCUERE")
'    Call BLANARCH("!SELCUERE")
'    '
'301 RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
'    If Len(RFF$) < 4 Then
'      Call MENSERR(24, "Sin Movimientos")
'      GoTo 10
'    End If
'    '
'    JJ& = 0
'    For UL& = 1 To Len(RFF$) Step 4
'      U& = CVS(Mid$(RFF$, UL&, 4))
'      X$ = REGLEIDO("COBRAN", U&)
'      NROLL& = CVS(Mid$(X$, 3, 4))
'      IMPO# = CVD(Mid$(X$, 15, 8))
'      FC% = CVI(Mid$(X$, 7, 2))
'      If FC% >= FLIXX% Then
'        If NROLL& > 0 Then
'          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
'          NNXX# = CDbl(NROLL&)
'          DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
'          Y$ = REGBLAN$("!SELCUERE")
'          Call REPLA(Y$, MKS$(U&), 1, 4)
'          Call REPLA(Y$, DOCU$, 5, 18)
'          Call REPLA(Y$, IMPOS$, 23, 13)
'          JJ& = JJ& + 1
'          Call GRAREG("!SELCUERE", Y$, JJ&)
'        End If
'      End If
'    Next UL&
'    Call CIERRARCH("!SELCUERE")
'    Call FRESHECHO("!SELCUERE")
'    Screen.MousePointer = 1
'    '
'    If JJ& < 1 Then
'      Call MENSERR(24, "Sin Recibos Posteriores al " + FENUE$)
'      GoTo 10
'    End If
'    '
'350 Call SELECHO("!SELCUERE")
'    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
'    If XVALO(REGCUEC&) < 1 Then
'        GoTo 999
'    End If
'    '
'    XX0$ = REGLEIDO$("COBRAN", REGCUEC&)
'    NROLL& = CVS(Mid$(XX0$, 3, 4))
'    NCLIE% = CVI(Left$(XX0$, 2))
'    '
'    TIDOCUM$ = "Recibo Nro."
'    NUDOCU$ = " " + TRIM$(Str$(NROLL&))
'    RACLI$ = TRIM$(Left$(RASOCLI$(NCLIE%), 8))
'    '
'    Screen.MousePointer = 11
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        If InStr(XX$, NUDOCU$) > 4 Then
'          If InStr(XX$, RACLI$) > 16 Then
'            DOCUNU& = CVS(Left$(XX$, 4))
'            GoTo 5
'          End If
'        End If
'      End If
'    Next UI&
'    Screen.MousePointer = 1
'    Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
'    GoTo 10
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 10
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'96  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 96
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Visible = False
'      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 350
'    '
'999 Call CIERRARCH("*.*")
'    Call LIBARCH("*.*")
'    '
'    Screen.MousePointer = 1
'    '
End Sub
'\\\OT-05-0441-WAR-30/06/05///
Sub VAPLIREC()
    '
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
        '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND ( TipoMovim = 'RCOB' OR (TipoMovim = 'FVEN'AND OpciMovim = 'NC' ) ) "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Lar ASC "
    Set TOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM APLICOBRA "
    SQLX$ = SQLX$ + " WHERE NuClien = " & NC% & ""
    SQLX$ = SQLX$ + " ORDER BY CoDocAp_Lar ASC "
    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    With TOTMP
      On Error Resume Next
      .MoveFirst 'habia un movenext
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      CODVIE$ = ""
      Do While Not .EOF
        NROX$ = !CODICOM_LAR   ' !CoDocAp_Lar
        IMPO# = !IHABECOMP     ' 0
        Call REPLA(Y$, NROX$, 5, 18)
        If IMPO# >= 0.005 Then
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call REPLA(Y$, MKS(JJ&), 1, 4)
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC%)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      RESTO# = XVALO(Mid$(RESEL$, 23, 13))
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      SQLX$ = " SELECT * FROM APLICOBRA "
      SQLX$ = SQLX$ + " WHERE NuClien = " & NC% & ""
      SQLX$ = SQLX$ + " ORDER BY CoDocOr_Lar ASC "
      Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
      '
      JJ& = 0
      With APCOTMP    ' ApliCobra
        .FindFirst "CoDocAp_Lar= '" & TRIM(DAPLI$) & "'"
307     If .NoMatch = False Then
          Y$ = REGBLAN$("!SELDOCAP")
          NROY$ = !CoDocOr_Lar
          REFY$ = Left(!RefeApli, 18)
          IMPO# = !ImpoApli
            If IMPO# > RESTO# Then IMPO# = RESTO#
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(Y$, NROY$, 1, 18)
            Call REPLA(Y$, REFY$, 19, 24)
            Call REPLA(Y$, IMPOS$, 44, 13)
            JJ& = JJ& + 1
            Call GRAREG("!SELDOCAP", Y$, JJ&)
            RESTO# = RESTO# - IMPO#
          End If
          .FindNext "CoDocAp_Lar= '" & TRIM(DAPLI$) & "'"
          GoTo 307
        End If
      End With
      '
      If RESTO# >= 0.005 Then
        Y$ = REGBLAN$("!SELDOCAP")
        NROY$ = "Pendiente de Aplicación"
        IMPOS$ = NUMSTRI$(RESTO#, 13, 2, 1)
        Call REPLA(Y$, NROY$, 19, 24)
        Call REPLA(Y$, IMPOS$, 44, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELDOCAP", Y$, JJ&)
      End If
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

Sub VAPLIFAC()
      '
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' AND OpciMovim <> 'NC'"
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Lar ASC "
    Set TOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM APLICOBRA "
    SQLX$ = SQLX$ + " WHERE NuClien = " & NC% & ""
    SQLX$ = SQLX$ + " ORDER BY CoDocOr_LAR ASC "
    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
    With TOTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      '
      Do While Not .EOF
        NROX$ = !CODICOM_LAR
        IMPO# = TOTMP!IDEBECOMP
        '
        If IMPO# >= 0.005 Then
          Call REPLA(Y$, NROX$, 5, 18)
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call REPLA(Y$, MKS(JJ&), 1, 4)
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(RASOCLI$(NC%)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      RESTO# = XVALO(Mid$(RESEL$, 23, 13))
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + RASOCLI$(NC%)
      '
      JJ& = 0
      With APCOTMP
        .FindFirst "CoDocOR_LAR= '" & TRIM(DAPLI$) & "'"
307     If .NoMatch = False Then
          Y$ = REGBLAN$("!SELDOCAP")
          NROY$ = !CoDocAp_Lar
          REFY$ = Left(!RefeApli, 18)
          IMPO# = !ImpoApli
            If IMPO# > RESTO# Then IMPO# = RESTO#
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(Y$, NROY$, 1, 18)
            Call REPLA(Y$, REFY$, 19, 24)
            Call REPLA(Y$, IMPOS$, 44, 13)
            JJ& = JJ& + 1
            Call GRAREG("!SELDOCAP", Y$, JJ&)
            RESTO# = RESTO# - IMPO#
          End If
          .FindNext "CoDocOr_Lar= '" & TRIM(DAPLI$) & "'"
          GoTo 307
        End If
      End With
      '
      If RESTO# >= 0.005 Then
        Y$ = REGBLAN$("!SELDOCAP")
        NROY$ = "Pendiente de Aplicación"
        IMPOS$ = NUMSTRI$(RESTO#, 13, 2, 1)
        Call REPLA(Y$, NROY$, 19, 24)
        Call REPLA(Y$, IMPOS$, 44, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELDOCAP", Y$, JJ&)
      End If
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
Sub CONAPLI()
  '
  Call APERBASDAT("CABAN")
  SQLX1$ = " SELECT FechEmisi, CodiCom_Lar FROM CAJABANC "
  Set NROTMP = CueCorri.OpenRecordset(SQLX1$, 4)
  '
  SQLX$ = " SELECT * FROM APLICOBRA "
  Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  With DOCAPTMP
    '
    On Error Resume Next
    .MoveLast
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    Max& = .RecordCount
    JJ& = 0
    YY$ = REGBLAN("APLICOB")
    FEC1% = 0
    .MoveFirst
    Do While Not .EOF
      Call REPLA(YY$, MKI(!NuClien), 3, 2)
      Call REPLA(YY$, !CoDocAp_Lar, 5, 18)
      Call REPLA(YY$, MKD(!ImpoApli), 23, 8)
      Call REPLA(YY$, Left(!RefeApli, 24), 31, 24)
      Call REPLA(YY$, !CoDocOr_Lar, 55, 18)
       DOC1$ = !CoDocAp_Lar
       With NROTMP
         On Error Resume Next
         .FindFirst "CodiCom_Lar = '" & DOC1$ & "' "
         If .NoMatch = True Then
          On Error GoTo 0
          GoTo 25
         End If
         On Error GoTo 0
          FEC1% = CVINT(!FechEmisi)
       End With
25    Call REPLA(YY$, MKI(FEC1%), 1, 2)
      JJ& = JJ& + 1
      Call TRACE(24, JJ&, Max&)
      Call GRAREG("APLICOB", YY$, JJ&)
      .MoveNext
    Loop
    '
  End With
  Call SETULTREG("APLICOB", JJ&)
  Call CIERRARCH("*.*")
  '
End Sub
'\\\OT-05-0441-WAR-FIN///
'\\\OT-05-0441-WAR-FIN///
Sub LIAPLIFE()
'   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = Int(CDbl(CVDAT(DES%)))
   HASA = Int(CDbl(CVDAT(HAS%)))
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT NUCLIEN, IHabeComp, FechEmisi, CodiCom_Lar FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & " "
   SQLX$ = SQLX$ + " AND ( TipoMovim = 'RCOB' OR (TipoMovim = 'FVEN'AND OpciMovim = 'NC' ) ) "
   SQLX$ = SQLX$ + " Order by FechEmisi,CodiCom_Lar ASC "
   Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   SQLX$ = " SELECT * FROM APLICOBRA "
   Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   With NROTMP
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 90
     On Error GoTo 0
     '
     Do While .EOF = False
       FEDOC% = CVINT(!FechEmisi)
       NCLIE% = !NuClien
       DOCAPLI$ = !CODICOM_LAR
       IMPOTO# = !IHABECOMP
       RESTO# = IMPOTO#
       If IMPOTO# < 0.005 Then GoTo 26
       '
       PRIREN% = 1
       SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
       DOCAPTMP.FindFirst SQLX2$
36     If DOCAPTMP.NoMatch = False Then
         YY$ = REGBLAN("APLICO1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKI(NCLIE%), 21, 2)
         If PRIREN% = 1 Then
           Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
           PRIREN% = 0
         End If
         Call REPLA(YY$, DOCAPTMP!CoDocOr_Lar, 31, 18)
           REFEA$ = Left(DOCAPTMP!RefeApli, 24)
         Call REPLA(YY$, REFEA$, 49, 24)
           IAPLI# = DOCAPTMP!ImpoApli
           If IAPLI# > RESTO# Then IAPLI# = RESTO#
         Call REPLA(YY$, MKD(IAPLI#), 73, 8)
         RESTO# = RESTO# - IAPLI#
         If IAPLI# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("APLICO1", YY$, JJ&)
           CAREGRA% = CAREGRA% + 1
         End If
         DOCAPTMP.FindNext SQLX2$
         GoTo 36
       End If
       '
       ' ESTO ES PARA PREVER QUE ESE RECIBO AUN NO SE APLICó NADA
       If RESTO# >= 0.005 Then
         ' buscar si aun está pendiente de aplicacion
         REFEB$ = "*** Pend.Aplic.***"
         SaDeuPen.FindFirst "Nuclien = " & NCLIE% & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         If SaDeuPen.NoMatch = True Then
           REFEB$ = "Aplic.Automatica"
         End If
         '
         YY$ = REGBLAN("APLICO1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKI(NCLIE%), 21, 2)
         If PRIREN% = 1 Then
            Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
         End If
         Call REPLA(YY$, REFEB$, 31, 18)
           REFEA$ = "Pendiente Aplic."
         Call REPLA(YY$, REFEA$, 49, 24)
         Call REPLA(YY$, MKD(RESTO#), 73, 8)
         JJ& = JJ& + 1
         Call GRAREG("APLICO1", YY$, JJ&)
       End If
       '
26   .MoveNext
     Loop
     '
   End With
   '
90 Call SETULTREG("APLICO1", JJ&)
   Call HEADERS("APLICO1", "")
   Call HEADERS("APLICO1", "Corresponde a: " + TRIM$(PERIO$))
   Call FINAL("*LIAPLIC1")
   '
99 Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
   '
   MODOMONEI% = 0
   MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ = "MULTIMON" Then
      MODOMONEI% = 2
   End If
   '
End Sub

Sub LISRETIIBB()

    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    DEFE = Int(CDbl(CVDAT(DES%)))
    HAFE = Int(CDbl(CVDAT(HAS%)))

    Call APERBASDAT("CABAN")
    JJ& = 0
    For U& = 1 To ULTREG&("MECOBRA")
        XXY$ = REGLEIDO$("MECOBRA", U&)
        CODALFA$ = Mid$(XXY$, 57, 4)
        If CODALFA$ = "RTIB" Then
             CUENCONT$ = TRIM$(Mid$(XXY$, 45, 12))
             DESCUENT$ = TRIM$(Mid$(XXY$, 31, 14))
             CODCUENT% = CUECOINT%(CUENCONT$)
             SQLX$ = "SELECT * FROM CAJABANC "
             SQLX$ = SQLX$ + "where INT(CDBL(FechEmisi)) >= " & DEFE & " "
             SQLX$ = SQLX$ + "AND INT(CDBL(FechEmisi)) <= " & HAFE & " "
             SQLX$ = SQLX$ + "AND CueContaI = " & CODCUENT% & " "
             '
             Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
             '
             REBLA1$ = REGBLAN$("SURETIB")
             With CABATEMP
                On Error Resume Next
                .MoveFirst
                If Err > 0 Then
                  On Error GoTo 0
                 ' GoTo 99
                End If
                On Error GoTo 0
                Do While Not .EOF
                   ZZ$ = REBLA1$
                   FECHAEMI = Int(CDbl(!FechEmisi))     'fecha emision del documento
                   TYNDOC$ = !CODICOM_LAR               'Documento asociado
                   DENCUEN$ = RASOCLI(!NuClien)         'Razon social del cliente
                   NuCuit$ = CUITCLI(!NuClien)          'CUIT del Cliente
                   IMPPER# = !IDEBECOMP                 'importe de percecpcion
                   
                   CajaBanc.FindFirst "CodiCom_Lar='" & TYNDOC$ & "' AND TipoMovim='RCOB'"
                   IMPREC# = CajaBanc!ValAbsComp                'Importe del recibo
                   
                   NUMCUE$ = CODCUE$(CODCUENT%)         'Numero de cuenta contable
                   DESCUECONT$ = DENOCUE$(CODCUENT%)    'Descripcion de cuenta contable
                   '
                   Call REPLA(ZZ$, MKI$(CVINT(!FechEmisi)), 1, 2)
                   Call REPLA(ZZ$, TYNDOC$, 3, 18)
                   Call REPLA(ZZ$, DENCUEN$, 21, 30)
                   Call REPLA(ZZ$, NuCuit$, 51, 14)
                   Call REPLA(ZZ$, MKD$(IMPREC#), 79, 8)
                   Call REPLA(ZZ$, MKD$(IMPPER#), 87, 8)
                   Call REPLA(ZZ$, NUMCUE$, 95, 12)
                   Call REPLA(ZZ$, DESCUECONT$, 107, 50)
                   '
                   JJ& = JJ& + 1
                   Call GRAREG("SUPERIB", ZZ$, JJ&)
                .MoveNext
                Loop
             End With
            '
          Call SETULTREG("SUPERIB", JJ&)
          Call HEADERS("SUPERIB", "")
          Call HEADERS("SUPERIB", "Periodo: " + PERIO$)
         
        End If
    Next U&
     Call CONECRUN("*SURETIB", "")
99   Call CIERRARCH("SUPERIB")
End Sub

Sub ATRACOBRA()
10 Call SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   If NCL1% < 1 And NCL2% < 1 Then GoTo 99
   
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% < 1 Then GoTo 10
   DES1# = Int(CDbl(CVDAT(DES%)))
   HAS1# = Int(CDbl(CVDAT(HAS%)))
   '
   SQLX$ = "SELECT * FROM CAJABANC"
   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'RCOB'"
   SQLX$ = SQLX$ + " AND NuClien >= " & NCL1% & ""
   SQLX$ = SQLX$ + " AND NuClien <= " & NCL2% & ""
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))>= " & DES1# & " "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta))<= " & HAS1# & " "
   Set PROATRATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With PROATRATEMP
          On Error Resume Next
          .MoveFirst
          If Err Then
             On Error GoTo 0
             Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
             GoTo 99
          End If
          '
          Y$ = REGBLAN("ATRAPROM")
          '
          U& = 0
          Do While Not .EOF
            NCLIE% = !NuClien
            If VENDSEL% < 1 Or VENDSEL% = VENDCLI%(NCLIE%) Then
               CLIENTE% = !NuClien
               NRECIBBO$ = SAFETEXT$(!CodiCom_Cor)
               ATRACHEQQUE% = XVALO(!AtraCheque)
               ATRACCANCE% = XVALO(!AtraCance)
               ATRASSO% = XVALO(!AtraCheque) + XVALO(!AtraCance)
               Call REPLA(Y$, MKI$(CLIENTE%), 1, 2)
               Call REPLA(Y$, NRECIBBO$, 3, 12)
               Call REPLA(Y$, MKS$(ATRACHEQQUE%), 15, 4)
               Call REPLA(Y$, MKS$(ATRACCANCE%), 19, 4)
               Call REPLA(Y$, MKS$(ATRASSO%), 23, 4)
               U& = U& + 1
               Call GRAREG("ATRAPROM", Y$, U&)
            End If
            .MoveNext
          Loop
          Call SETULTREG("ATRAPROM", U&)
          Call CIERRARCH("ATRAPROM")
          Call FILESORT("ATRAPROM", "", 3, 3, "ASC")
          Call FILESORT("ATRAPROM", "", 1, 1, "ASC")
     End With
     NomV$ = DEVENDE$(VENDSEL%)
     Call HEADERS("ATRAPROM", "")
     Call HEADERS("ATRAPROM", "Periodo: " + PERIO$)
     Call HEADERS("ATRAPROM", "Rango de Cuentas: " + TRIM$(Str$(NCL1%)) + " - " + TRIM$(Str$(NCL2%)))
     If VENDSEL% > 0 Then Call HEADERS("ATRAPROM", "Vendedor: " + NomV$)

     Call FINAL("*ATRAPROM")
99   Screen.MousePointer = 1
   
End Sub

