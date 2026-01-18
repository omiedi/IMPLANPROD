VERSION 5.00
Begin VB.Form OPORPA00 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones Complementarias - Ordenes de Pago"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   855
      Left            =   4305
      Picture         =   "OPORPA00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Analisis de Pagos por Fecha Acreditación de Cheque"
      Top             =   3465
      Width           =   855
   End
   Begin VB.CommandButton Command8 
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
      Height          =   855
      Left            =   4305
      Picture         =   "OPORPA00.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4410
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Emision de Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   2
      Left            =   210
      TabIndex        =   21
      Top             =   3675
      Width           =   3900
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   22
         Top             =   420
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
            Picture         =   "OPORPA00.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Listado de Ordenes de Pago Por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Or- denes de Pago"
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
            TabIndex        =   24
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4305
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   5460
      Width           =   855
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
      Left            =   4305
      Picture         =   "OPORPA00.frx":075E
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
      Left            =   4305
      Picture         =   "OPORPA00.frx":08A8
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
      Height          =   1275
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   5040
      Width           =   3900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
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
            Picture         =   "OPORPA00.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tabla de Me- dios de Pago"
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
            TabIndex        =   10
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
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
      Height          =   3375
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3900
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   16
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
            Picture         =   "OPORPA00.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ajuste Directo de Saldos Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Sal- dos Acreedores"
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
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   1995
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
            Picture         =   "OPORPA00.frx":0E46
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Imputación de Créditos y Anticipos a Proveedores"
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
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   80
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
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
            Picture         =   "OPORPA00.frx":0F90
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Orden de Pago ya Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver por Pantalla y Repetir Impresion"
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
            TabIndex        =   3
            Top             =   75
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   25
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
            Picture         =   "OPORPA00.frx":15FA
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Imprimir Copia de Orden de Pago ya Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Des-Aplicar Ordenes de Pago"
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
            TabIndex        =   27
            Top             =   75
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
            Picture         =   "OPORPA00.frx":1744
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de Orden de Pago Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Orden de Pago"
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
            TabIndex        =   6
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   75
            Width           =   1530
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3730
      Picture         =   "OPORPA00.frx":1A4E
      Top             =   5780
      Width           =   2010
   End
End
Attribute VB_Name = "OPORPA00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If (USNBR% Mod 100) <> 1 Then
       Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
       Command1.Enabled = True
       Exit Sub
    End If
    '
    LU$ = LUDAT$
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    If COPYFILE%(LU$ + "MEPAGO.DAT", "C:\FLEXOFT\MEPAGO.WKF") <> 1 Then
        Command1.Enabled = True
        Exit Sub
    End If
    '
    Dim REMEPA&(1000)
    '
360 ATY% = VENTABU%("COMEPAG/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MEPAGO")
       X$ = REGLEIDO$("!MEPAGO", U&)
       CDX% = CVI(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             '
             If UCase$(Mid$(X$, 57, 4)) = "CH.3" Then
                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             If UCase$(Mid$(X$, 57, 4)) = "PROP" Then
                 Call MENSERR(24, "Valor '" + Mid$(X$, 57, 4) + "' no Permitido")
                 ATX% = 1
             End If
             '
             J& = J& + 1
363          Call GRAREG("!MEPAGO", X$, J&)
             If TRIM$(Mid$(X$, 3, 12)) = "" Then ATX% = 1
             If ECOARCH$("ECUECON", TRIM$(Mid$(X$, 45, 12))) = "" Then ATX% = 1
             CCII% = CUEINT%(TRIM$(Mid$(X$, 45, 12)))
             Call REPLA(X$, MKI$(CCII%), 61, 2)
             If CCII% < 1 Or (NUCUEN% > 0 And CCII% > NUCUEN%) Then ATX% = 1
364          If TRIM$(Mid$(X$, 57, 4)) = "" Then ATX% = 1
             '
             If REMEPA&(CDX%) < 1 Then REMEPA&(CDX%) = J&
             '
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MEPAGO", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        For UII% = 0 To 1000: REMEPA&(UII%) = 0: Next UII%
        Call COMUNI("Datos Incompletos o no Válidos.")
        GoTo 360
    End If
    '
    JJ& = 0
    For UII% = 1 To 1000
      If REMEPA&(UII%) > 0 Then
        If REMEPA&(UII%) <= J& Then
          X$ = REGLEIDO$("!MEPAGO", REMEPA&(UII%))
          JJ& = JJ& + 1
          Call GRAREG("MEPAGO", X$, JJ&)
        End If
      End If
    Next UII%
    Call SETULTREG("MEPAGO", JJ&)
    Call CIERRARCH("MEPAGO")
    Call CIERRARCH("!MEPAGO")
    Call FRESHECHO("MEPAGO")
    '
390 Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    '
    Hide
    '
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
    Screen.MousePointer = 11
    '
    If EXISTE%(LUDAT$ + "INDIORPA.RFS") < 1 Then
      Call COPYSTRU("INDIPRES", "INDIORPA")
    End If
    '
    MAXDEU& = ULTREG&("CUECORR")
    For REG& = 1 To MAXDEU&
      X$ = REGLEIDO$("CUECORR", REG&)
      NCLIE% = CVI(Mid$(X$, 5, 2))
      TIDOC$ = UCase$(Mid$(X$, 7, 2))
      If TIDOC$ = "OP" Then
        IMPO# = Abs(CVD(Mid$(X$, 115, 8))) + Abs(CVD(Mid$(X$, 123, 8)))
        If IMPO# >= 0.005 Then
          IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
          NPX& = Val(Mid$(X$, 17, 8))
          FEX = CVI(Mid$(X$, 1, 2))
          YY$ = MKS$(NPX&) + TRIM$(RASOCLI$((-1) * NCLIE%)) + " - " + FECHATEX$(FEX)
          JJ& = JJ& + 1
          Call GRAREG("INDIORPA", YY$, JJ&)
        End If
      End If
    Next REG&
    Call SETULTREG("INDIORPA", JJ&)
    Screen.MousePointer = 1
    '
3   Call SELECHO("INDIORPA/Indice de Ordenes de Pago Emitidas")
    NPX& = Val(VALACT1$("INDIORPA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Pago"
    NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
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
    GoTo 99
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
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command23.Enabled = True
    '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call ANUORPA         ' Anulación de Orden de Pago
    Command25.Enabled = True
End Sub

Private Sub Command3_Click()
    Hide
    Call IMPUPRO1
    On Error Resume Next
    FIMPUPRO.Hide
    On Error GoTo 0
    FIMPUPRO.Show 1
    Call IMPUPRO2
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call OPNOIN("Ajuste de Saldos de Proveedores")
    Call AJUSALP
    Screen.MousePointer = 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    '
10  FEX = HOY(HO$)
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
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
    End If
    '
    CANOPFF% = 0
    FIN& = ULTREG&("CUECORR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If Mid$(CC0$, 7, 2) = "FF" Then
        If CVI(Left$(CC0$, 2)) >= DES% Then
          If CVI(Left$(CC0$, 2)) <= HAS% Then
            NOPA = Val(Mid$(CC0$, 20, 6))
            If NOPA <> NOPFF(CANOPFF%) Then
               CANOPFF% = CANOPFF% + 1
               NOPFF(CANOPFF%) = NOPA
            End If
          End If
        End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      FEOPA% = CVI(CH8$)
      If FEOPA% >= DES% Then                ' CORRESPONDE FECHA
        If FEOPA% <= HAS% Then
          IM# = CVD(CH5$)
          If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHEQUEMI", J&)
                  CH1$ = Mid$(CH0$, 1, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 41, 4)
                  If CVI(CH1$) = NP% Then
                     If CVS(CH6$) = NOPA Then
                        ITOT# = ITOT# + CVD(CH5$)
                     End If
                  End If
               Next J&
               '
               If TICUEPRO((-1) * NP%) = 1 Then ' CUENTA DE GASTOS
                  For UDP& = 1 To ULTREG&("DETPAGAS")
                     DTP$ = REGLEIDO$("DETPAGAS", UDP&)
                     If CVS(Mid$(DTP$, 3, 4)) = NOPA Then
                        CUEDP% = CVI(Mid$(DTP$, 17, 2))
                        If TRIM$(CODCUE$(CUEDP%)) <> "" Then
                           IDP# = CVD(Mid$(DTP$, 9, 8))
                           Z$ = REGBLAN$("LISORPA")
                           Call REPLA(Z$, MKI$(FEOPA%), 1, 2)
                           Call REPLA(Z$, MKS$(NOPA), 3, 4)
                           Call REPLA(Z$, MKI$(NP%), 7, 2)
                           Call REPLA(Z$, CODCUE$(CUEDP%), 9, 16)
                           Call REPLA(Z$, RASOCLI$((-1) * NP%), 25, 24)
                           Call REPLA(Z$, MKD$(IDP#), 49, 8)
                           JJ& = JJ& + 1
                           Call GRAREG("LISORPA", Z$, JJ&)
                           ITOT# = ITOT# - IDP#
                        End If
                     End If
                  Next UDP&
               End If
               '
               If Abs(ITOT#) >= 0.005 Then
                 Z$ = REGBLAN$("LISORPA")
                 Call REPLA(Z$, MKI$(FEOPA%), 1, 2)
                 Call REPLA(Z$, MKS$(NOPA), 3, 4)
                 Call REPLA(Z$, MKI$(NP%), 7, 2)
                 Call REPLA(Z$, CUEMADRE$((-1) * NP%), 9, 16)
                 Call REPLA(Z$, RASOCLI$((-1) * NP%), 25, 24)
                 Call REPLA(Z$, MKD$(ITOT#), 49, 8)
                 JJ& = JJ& + 1
                 Call GRAREG("LISORPA", Z$, JJ&)
               End If
               '
               NOPANT = NOPA
               NPANT% = NP%
            End If
            '
            CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
            CH1$ = Mid$(CH0$, 1, 2)
            CH2$ = Mid$(CH0$, 3, 12)
            CH3$ = Mid$(CH0$, 15, 16)
            CH5$ = Mid$(CH0$, 33, 8)
            CH6$ = Mid$(CH0$, 41, 4)
            CH7$ = Mid$(CH0$, 45, 2)
            CH8$ = Mid$(CH0$, 47, 2)
            '
            CHE$ = ""
            For JI% = 1 To Len(CH2$)
               If Asc(Mid$(CH2$, JI%)) > 32 Then CHE$ = CHE$ + Mid$(CH2$, JI%, 1)
            Next JI%
            '
            NCHE$ = ""
            For JI% = 1 To Len(CH3$)
               If Asc(Mid$(CH3$, JI%)) > 32 Then NCHE$ = NCHE$ + Mid$(CH3$, JI%, 1)
            Next JI%
            '
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 18
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
18          CUEH% = CUEMP%(UX%)
            IPASE# = CVD(CH5$)
            DTX$ = TRIM$(CHE$)
            If DTX$ <> "" Then DTX$ = DTX$ + " - "
            DTX$ = DTX + TRIM$(NCHE$)
            '
            Z$ = REGBLAN$("LISORPA")
            Call REPLA(Z$, MKI$(FEOPA%), 1, 2)
            Call REPLA(Z$, MKS$(NOPA), 3, 4)
            Call REPLA(Z$, MKI$(NP%), 7, 2)
            Call REPLA(Z$, CODCUE$(CUEH%), 9, 16)
            Call REPLA(Z$, DTX$, 25, 24)
            Call REPLA(Z$, MKD$(IPASE#), 57, 8)
            JJ& = JJ& + 1
            Call GRAREG("LISORPA", Z$, JJ&)
            '
          End If
        End If
      End If
76 Next IAX&
   '
   Call SETULTREG("LISORPA", JJ&)
   Call CIERRARCH("LISORPA")
   Screen.MousePointer = 1
   '
   Call FINAL("*SORORPA")
99 Screen.MousePointer = 1
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call DESAPLIOPA         ' Des-aplicacion de Orden de Pago
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Call HELPONLINE("ORPASTA")
End Sub

Sub DESAPLIOPA()       ' DES-APLICACION de ordenes de pago
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
     '
     FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
     '
10   Call SELECHO("ACREDOR")
     NC% = Val(VALACT1$("ACREDOR"))
     If NC% < 1 Then Exit Sub
     '
100  Call BLANARCH("!SELCUECO")
     MAXDEU& = ULTREG&("CUECORR")
     RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
     '
     For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(X$, 5, 2)) <> NC% Then
            X$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", X$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        TIDOC$ = UCase$(Mid$(X$, 7, 2))
        If TIDOC$ = "OP" Then
           IMPO# = Abs(CVD(Mid$(X$, 115, 8))) + Abs(CVD(Mid$(X$, 123, 8)))
           If IMPO# >= 0.005 Then
              IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
              Y$ = MKS$(REG&) + Mid$(X$, 7, 18) + IMPOS$
              Call REGAPP("!SELCUECO", Y$)
           End If
        End If
     Next U&
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Ordenes de Pago que Puedan Desaplicarse.")
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
     XX0$ = REGLEIDO$("CUECORR", REGCUEC&)
     IMPO# = Abs(CVD(Mid$(XX0$, 115, 8))) + Abs(CVD(Mid$(XX0$, 123, 8)))
     IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     X$ = TRIM$(Mid$(XX0$, 7, 18))
     Y$ = IMPOS$
     DOCAPLI$ = X$
     '
     If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Abandona\Confirma Desaplicación") <> 2 Then GoTo 10
     '
     ' DESAPLICACION
     ' *************
     '
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("APLIPAG")
       X$ = REGLEIDO$("APLIPAG", U&)
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
           Call REAPLICREPRO(NC%)
        End If
     End If
     '
     GoTo 10
     '
End Sub
'

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    Dim NORP&(8192), NPRO%(8192), FOPA%(8192), IOPA#(8192)
    '
10  FEX = HOY(HO$)
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
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
    End If
    '
    CANOPFF% = 0
    CANOPAI% = 0
    FIN& = ULTREG&("CUECORR")
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If CVI(Left$(CC0$, 2)) >= DES% Then
        If CVI(Left$(CC0$, 2)) <= HAS% Then
          If UCase$(Mid$(CC0$, 12, 4)) <> "ANUL" Then
            NOPA& = Val(Mid$(CC0$, 18, 8))
            PROV% = CVI(Mid$(CC0$, 5, 2))
            FEMI% = CVI(Mid$(CC0$, 1, 2))
            IMPX# = CVD(Mid$(CC0$, 115, 8))
            If Mid$(CC0$, 7, 2) = "FF" Then
                If NOPA <> NOPFF(CANOPFF%) Then
                  CANOPFF% = CANOPFF% + 1
                  NOPFF(CANOPFF%) = NOPA
                End If
              ElseIf Mid$(CC0$, 7, 2) = "OP" Then
                CANOPAI% = CANOPAI% + 1
                NORP&(CANOPAI%) = NOPA&
                NPRO%(CANOPAI%) = PROV%
                FOPA%(CANOPAI%) = FEMI%
                IOPA#(CANOPAI%) = IMPX#
            End If
          End If
        End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0: UENCU% = 1
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      FEOPA% = CVI(CH8$)
      If FEOPA% >= DES% Then                ' CORRESPONDE FECHA
        If FEOPA% <= HAS% Then
          IM# = CVD(CH5$)
          If Abs(IM#) >= 0.005 Then
            NOPA& = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            '
            vuelta% = 0
20          For KKI% = UENCU% To CANOPAI%
              If NOPA& = NORP&(KKI%) Then
                If NP% = NPRO%(KKI%) Then
                  If FEOPA% = FOPA%(KKI%) Then
                    IOPA#(KKI%) = IOPA#(KKI%) - IM#
                    UENCU% = KKI%
                    GoTo 76
                  End If
                End If
              End If
            Next KKI%
            If vuelta% = 0 Then
              vuelta% = 1
              UENCU% = 1
              GoTo 20
            End If
            CANOPAI% = CANOPAI% + 1
            NORP&(CANOPAI%) = NOPA&
            NPRO%(CANOPAI%) = NP%
            FOPA%(CANOPAI%) = FEOPA%
            IOPA#(CANOPAI%) = 0 - IM#
            '
          End If
        End If
      End If
76 Next IAX&
   '
80 For KKI% = 1 To CANOPAI%
     If Abs(IOPA#(KKI%)) >= 0.005 Then
       NOPA& = NORP&(KKI%)
       NP% = NPRO%(KKI%)
       FEX = FOPA%(KKI%)
       Call MENSERR(24, "Error en Orden de Pago" + Str$(NOPA&) + "\de " + TRIM$(RASOCLI$((-1) * NP%)) + "\por " + TRIM$(CSTRING$(MKD$(IOPA#(KKI%)), 4, 12, 2, 2)) + " del " + FECHATEX$(FEX))
     End If
   Next KKI%
   '
99 Command8.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("ANAPAGOS", "Analisis de Pagos por Cheque")
   Command9.Enabled = True
End Sub

Sub ANUORPA()       ' anulacion de ordenes de pago
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
     '
     FECHALIM% = Val(CONTROL$("", "CIEMECOM"))
     '
10   Call SELECHO("PROVED1")
     NC% = Val(VALACT1$("PROVED1"))
     If NC% < 1 Then Exit Sub
     '
100  Call BLANARCH("!SELCUECO")
     MAXDEU& = ULTREG&("CUECORR")
     RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
     '
     For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(X$, 5, 2)) <> NC% Then
            X$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", X$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        If CVI(Left$(X$, 2)) > FECHALIM% Then
           TIDOC$ = UCase$(Mid$(X$, 7, 2))
           If TIDOC$ = "OP" Then
              IMPO# = Abs(CVD(Mid$(X$, 115, 8))) + Abs(CVD(Mid$(X$, 123, 8)))
              If IMPO# >= 0.005 Then
                 IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
                 Y$ = MKS$(REG&) + Mid$(X$, 7, 18) + IMPOS$
                 Call REGAPP("!SELCUECO", Y$)
              End If
           End If
        End If
     Next U&
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Ordenes de Pago que Puedan Anularse.")
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
     XX0$ = REGLEIDO$("CUECORR", REGCUEC&)
     IMPO# = Abs(CVD(Mid$(XX0$, 115, 8))) + Abs(CVD(Mid$(XX0$, 123, 8)))
     IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     X$ = TRIM$(Mid$(XX0$, 7, 18))
     Y$ = IMPOS$
     '
     If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Abandona\Confirma Anulacion") <> 2 Then GoTo 10
     '
     ' anulacion
     ' *********
     '
3011 Call BLOQARCH("CUECORR")
     Call BLOQARCH("PROVED2")
     Call BLOQARCH("DETPAGAS")
     Call BLOQARCH("CAJABANC")
     Screen.MousePointer = 11
     '
     NORPEX$ = Left$(X$, 3) + Mid$(X$, 13, 6) + "-" + Mid$(X$, 8, 2)
     '
     FF1% = CVI(Left$(XX0$, 2))
3012 NF$ = Mid$(XX0$, 7, 18)
     Call REPLA(NF$, "anul", 6, 4)
     NORPA& = Val(Mid$(NF$, 11, 8))
     PROV% = CVI(Mid$(XX0$, 5, 2))
3013 XD# = CVD(Mid$(XX0$, 115, 8))
3014 XH# = CVD(Mid$(XX0$, 123, 8))
3015 SS8X$ = Mid$(XX0$, 139, 60)
     IM# = CVD(Mid$(SS8X$, 3, 8)) + XD# - XH#
     Call REPLA(SS8X$, MKI$(FF1%), 1, 2)
     Call REPLA(SS8X$, MKD$(IM#), 3, 8)
     Call REPLA(XX0$, NF$, 7, 18)
     Call REPLA(XX0$, MKD$(0), 115, 8)
     Call REPLA(XX0$, MKD$(0), 123, 8)
     Call REPLA(XX0$, SS8X$, 139, 60)
     Call GRAREG("CUECORR", XX0$, REGCUEC&)
     '
     NC1% = (-1) * NC%
     RC& = REGICLI&(NC1%)
     DX$ = REGLEIDO$("PROVED2", RC&)
     SALDO# = CVD(Mid$(DX$, 109, 8)) + XD# - XH#
     Call REPLA(DX$, MKD$(SALDO#), 109, 8)
     Call GRAREG("PROVED2", DX$, RC&)
     '
     FIN& = ULTREG&("CAJABANC")
     '
     For U& = FIN& To 1 Step -1
        Call TRACE(20, 1 + FIN& - U&, FIN&)
        Z$ = REGLEIDO$("CAJABANC", U&)
        If Left$(Z$, 12) = NORPEX$ Then
          Call REPLA(Z$, Chr$(9), 121, 1)
          Call REPLA(Z$, IDACTU$, 269, 6)
          Call GRAREG("CAJABANC", Z$, U&)
        End If
        '
        If Mid$(Z$, 277, 12) = NORPEX$ Then
          If Asc(Mid$(Z$, 121, 1)) = 2 Then ' cheque salido de cartera
            Call REPLA(Z$, Chr$(1), 121, 1)  ' Marca Cheque en Cartera
          End If
          Call REPLA(Z$, IDACTU$, 263, 6)  ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
          Call REPLA(Z$, MKI$(0), 275, 2)
          Call REPLA(Z$, Space$(166), 277, 166)
          Call GRAREG("CAJABANC", Z$, U&)
        End If
     Next U&
     Call CIERRARCH("CAJABANC")
     '
     FIN& = ULTREG&("DETPAGAS")
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        Z$ = REGLEIDO$("DETPAGAS", U&)
        If CVS(Mid$(Z$, 3, 4)) = NOP& Then
          If CVI(Mid$(Z$, 7, 2)) = FF1% Then
            If CVI(Left$(Z$, 2)) = PROV% Then
              Call REPLA(Z$, MKD$(0), 9, 8)
              Call GRAREG("DETPAGAS", Z$, U&)
            End If
          End If
        End If
     Next U&
     Call CIERRARCH("DETPAGAS")
     Screen.MousePointer = 1
     '
     If TICUEPRO%((-1) * NC%) = 0 Then
        TTXX$ = "Anulación Completa.\  \Desea Reconstruir la Composición de Deuda\de la Cuenta '" + TRIM$(RASOCLI$((-1) * NC%)) + "' ?"
        If COMALTER%(TTXX$ + "\\Reconstruir\Continuar") = 1 Then
           Call REAPLICREPRO(NC%)
        End If
     End If
     '
     Call CIERRARCH("CUECORR")
     Call CIERRARCH("PROVED2")
     '
     GoTo 10
     '
End Sub
'
Private Sub Form_Load()
   If EXISTE%("ANAPAGOS.EXE") < 1 Then
     Command9.Visible = False
   End If
End Sub

