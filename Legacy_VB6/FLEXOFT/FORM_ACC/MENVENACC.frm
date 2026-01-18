VERSION 5.00
Begin VB.Form MENVENACC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Ventas - Cuentas Corrientes"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5205
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   4320
      ScaleHeight     =   4515
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command4 
         Caption         =   "Proce"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   2950
         Visible         =   0   'False
         Width           =   600
      End
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
         Height          =   600
         Left            =   120
         Picture         =   "MENVENACC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Sub-Sistema de Presupuestos y Pedidos de Clientes"
         Top             =   2280
         Width           =   600
      End
      Begin VB.CommandButton Command29 
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
         Left            =   120
         Picture         =   "MENVENACC.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   900
         Width           =   600
      End
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
         Left            =   120
         Picture         =   "MENVENACC.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   300
         Width           =   600
      End
      Begin VB.CommandButton Command18 
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
         Left            =   120
         Picture         =   "MENVENACC.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Acceso a Listas de Precios de Venta Y Condiciones Comerciales"
         Top             =   1680
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3240
         Width           =   600
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
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MENVENACC.frx":0BA0
         Top             =   3480
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0F0C0&
      Caption         =   "Gestión de Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   4065
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Control de Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   3585
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   420
            Width           =   2325
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
               Picture         =   "MENVENACC.frx":2AC2
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Acceso al Control General de Gestión Comercial"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión Comercial"
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
               TabIndex        =   11
               Top             =   105
               Width           =   1620
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0F0C0&
         Caption         =   "Operaciones Comerciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   4
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command14 
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
               Picture         =   "MENVENACC.frx":2DCC
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir o Registrar un Recibo de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Recibo Cobranza en Cta. Corriente"
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
               TabIndex        =   5
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   6
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Picture         =   "MENVENACC.frx":2F16
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar una Factura en Cuenta Corriente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación C.Cte"
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
               TabIndex        =   7
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op_comerc 
      Caption         =   "&Operaciones Comerciales"
      Begin VB.Menu fac_ven 
         Caption         =   "Facturación de Ventas"
      End
      Begin VB.Menu recibo 
         Caption         =   "Recibo de Cobranza"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "&Consultas e Informes"
      Begin VB.Menu cta_cte 
         Caption         =   "Resumen de Cuenta Corriente"
         Begin VB.Menu scr1 
            Caption         =   "Individual por Pantalla"
         End
         Begin VB.Menu print1 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu fac_pen 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu plh20 
         Caption         =   "-"
      End
      Begin VB.Menu avan_fac 
         Caption         =   "Avance de Facturación"
         Begin VB.Menu avan_fac1 
            Caption         =   "Facturas, N.Debito y N.Credito"
            Begin VB.Menu scr3 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print3 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu avan_fac2 
            Caption         =   "Cobranza (Cuenta Corriente)"
            Begin VB.Menu scr4 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print4 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu res_mens 
         Caption         =   "Resumenes Mensuales"
         Begin VB.Menu sub_iva 
            Caption         =   "Subdiario de IVA Ventas"
         End
         Begin VB.Menu sub_ven 
            Caption         =   "Subdiario de Ventas"
         End
         Begin VB.Menu iva 
            Caption         =   "-"
         End
         Begin VB.Menu cons_tipocom 
            Caption         =   "Consulta por Tipo de Comprobante"
            Begin VB.Menu ctipo_fac 
               Caption         =   "Facturas"
               Begin VB.Menu scr5 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print5 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_ndeb 
               Caption         =   "Notas de Debito"
               Begin VB.Menu scr6 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print6 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_ncre 
               Caption         =   "Notas de Credito"
               Begin VB.Menu scr7 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print7 
                  Caption         =   "Imprimir / Exportar"
               End
            End
            Begin VB.Menu ctipo_reci 
               Caption         =   "Recibos de Cobranza"
               Begin VB.Menu scr9 
                  Caption         =   "Por Pantalla"
               End
               Begin VB.Menu print9 
                  Caption         =   "Imprimir / Exportar"
               End
            End
         End
         Begin VB.Menu plh21 
            Caption         =   "-"
         End
         Begin VB.Menu resmen_cta 
            Caption         =   "Resumen Mensual por Cuenta"
            Begin VB.Menu scr8 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print8 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu otros 
         Caption         =   "Otros"
         Begin VB.Menu lim_cre 
            Caption         =   "Deudas y Limite de Crédito"
            Begin VB.Menu scr2 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu print2 
               Caption         =   "Listado Impreso"
            End
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "MENVENACC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    OPP% = COMALTER%("Gestión de Ventas y Deudores\\Minimizar\Cerrar\Cancelar")
    If OPP% = 1 Then
        WindowState = 1
      ElseIf OPP% = 2 Then
        Call FINAL("")
        Exit Sub
    End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("RECOBRAN", "Recibo de Cobranza en Cuenta Corriente") >= 2 Then
      TC = TICAMONE(-1) ' fuerza refresco de tabla
      RECIBO04.Show 1
   End If
   Command14.Enabled = True
End Sub


Private Sub Command18_Click()
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.\Descuentos Combinados")
    If MODX% = 1 Then
         Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTASIVA
         Call COMUNI("Re-Inicie Ahora la Aplicación.")
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf MODX% = 4 Then
         DEFDESCO.Show 1
    End If
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   OPX% = ALTERNA%("Presupuestos a Clientes\Pedidos de Clientes")
   If OPX% = 1 Then
        Call CONECRUN("PRESUP04", "Presupuestos a Clientes")
      ElseIf OPX% = 2 Then
        Call CONECRUN("PEDIDO04", "Pedidos de Clientes")
   End If
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    '
    'SI ES FACTURA ELECTRONICA NO REALIZA VALIDACION DE CAI
    '******************************************************
    If DERACCE%("FACTACTE", "Facturación en Cuenta Corriente") >= 2 Then
      If CONTROL$("FACTURA", "SALIDA") <> "FELECT" Then

        MODU$ = "FACTURACION DE VENTAS"
        HOII% = HOY(HOS$)
        If EXISTE%(LUDAT$ + "VENCICAI.RFS") = 1 Then
          For UU& = 1 To ULTREG&("VENCICAI")
            XYZ$ = REGLEIDO$("VENCICAI", UU&)
            FEVENCAI% = CVI(Mid$(XYZ$, 63, 2))
            If FEVENCAI% > 0 Then
              If FEVENCAI% < HOII% Then
                DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                   Call MENSERR(24, "Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
                'Call MENSERR(24, "Imposible Acceder a Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
                'GoTo 99
              ElseIf FEVENCAI% < HOII% + 32 Then
                    DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                    Call MENSERR(24, "ATENCION: \CAI de " + DEFORMU$ + "\Próximo a Vencer !!!")
              End If
            End If
          Next UU&
        '
        End If
      End If
      '
      TC = TICAMONE(-1) ' fuerza refresco de tabla
      '
      EPAC$ = EMPREAC$
      If InStr(UCase$(EPAC$), "GRUPO BRASIL") > 0 Then
         OPX% = COMALTER%("Opciones de Trabajo para Facturación:\\Normal Inter-Activa\Automática")
         If OPX% <> 2 Then GoTo 98
         '
         Call FACABONOS
         GoTo 99
      End If
      '
98    FORFACCS.Show 1
    End If
99  Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   COGEVE04.Show 1
   Command3.Enabled = True
End Sub


Private Sub Command4_Click()
      Command4.Enabled = False
      Call ABRECABAN
      If CLACONTRA% <> 1 Then GoTo 99
      
      Call DESHASFECHA(DES%, HAS%, PERIO$)
      If PERIO$ = "" Then GoTo 99

      DESDEX = CDbl(CVDAT(DES%))
      HASTAX = CDbl(CVDAT(HAS%))
      
      SQLX$ = "SELECT COUNT(*) AS TOTAL FROM CAJABANC "
      SQLX$ = SQLX$ + " WHERE CODICOM_LAR LIKE 'RB*' AND TIPOMOVIM = 'CHRECI'  AND CUECONTAI = 34"
      SQLX$ = SQLX$ + " AND FECHEMISI >= " & DESDEX & " "
      SQLX$ = SQLX$ + " AND FECHEMISI <= " & HASTAX & " "
      Dim RST As Recordset
      Set RST = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      TOT& = RST!Total
      
      Call COMUNI("Se van a Modificar: " & TOT& & " Registros")
      
      SQLX$ = "update  cajabanc set cuecontai = 6"
      SQLX$ = SQLX$ + " WHERE CODICOM_LAR LIKE 'RB*' AND TIPOMOVIM = 'CHRECI'  AND CUECONTAI = 34"
      SQLX$ = SQLX$ + " AND FECHEMISI >= " & DESDEX & " "
      SQLX$ = SQLX$ + " AND FECHEMISI <= " & HASTAX & " "
    
      CueCorri.Execute SQLX$
      '
      SQLX$ = "SELECT COUNT(*) AS TOTAL FROM CAJABANC "
      SQLX$ = SQLX$ + " WHERE CODICOM_LAR LIKE 'RB*' AND TIPOMOVIM = 'CHRECI'  AND CUECONTAI = 34"
      SQLX$ = SQLX$ + " AND FECHEMISI >= " & DESDEX & " "
      SQLX$ = SQLX$ + " AND FECHEMISI <= " & HASTAX & " "
      Dim RST1 As Recordset
      Set RST1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
      TOT& = RST1!Total
      If TOT& < 1 Then
        Call COMUNI("Proceso Finalizado Correctamente")
        GoTo 99
      End If
      Call COMUNI("No se Han Podido Modificar Los Registros")
      '
99    Command4.Enabled = True
      
End Sub

Private Sub exit_Click()
   Call Command1_Click
End Sub

Private Sub fac_pen_Click()
   Call COGEVE04.Command3_Click
End Sub

Private Sub fac_ven_Click()
   Call Command22_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   HOII% = HOY(HOS$)
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "RECIBO" Then
       RECIBO04.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 7) = "FACTURA" Then
       FORFACCS.Show 1
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "LPRECIOS" Then
       Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "CGVENTAS" Then
       COGEVE04.Show 1
       Call FINAL("")
   End If
   '
   DoEvents
   CUEM$ = CONTROL$("SVD451S", "CUEIVA")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   CUEM$ = CONTROL$("SVD451S", "CUEADU")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   CUEM$ = CONTROL$("SVD451S", "CUERET")
   If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   '
   For IXX% = 0 To 11
     CUEM$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(IXX%)))
     If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then GoTo 50
   Next IXX%
   '
   Screen.MousePointer = 1
   GoTo 99
   '
50 Call MENSERR(24, "Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal.")
   'Call CONECRUN("FSETUP04/OPCLAVEN", "Configuración de Funcionamiento")
   OPCLAV04.Show 1
   Call FINAL("")
   '
99 End Sub

Sub RECUREGALIA()
   '
   AAX$ = LOADFILE$("FLOGON.EXE")
   TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
   Mid$(AAX$, 1, 128) = TTXX$
   '
   BBX$ = Space$(128)
   Call REPLA(BBX$, "01", 1, 2)
   Call REPLA(BBX$, "SALDO JUNIO / JULIO", 3, 20)
   Call REPLA(BBX$, "05/08/07", 23, 8)
   Call REPLA(BBX$, "$", 31, 4)
   Call REPLA(BBX$, "      357.35", 35, 12)
   SUCAR& = 0: For I% = 1 To 120
                  SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
               Next I%
   Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
   CCX$ = ENCRIPTA$(BBX$)
   Mid$(AAX$, 129, 128) = CCX$
   '
   BBX$ = Space$(128)
   Call REPLA(BBX$, "02", 1, 2)
   Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
   Call REPLA(BBX$, "05/08/07", 23, 8)
   Call REPLA(BBX$, "$", 31, 4)
   Call REPLA(BBX$, "      220.00", 35, 12)
   SUCAR& = 0: For I% = 1 To 120
                  SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
               Next I%
   Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
   CCX$ = ENCRIPTA$(BBX$)
   Mid$(AAX$, 257, 128) = CCX$
   '
   Call SAVEFILE("MSTDRV04.EXE", AAX$)
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub print1_Click()
   RECUEC04.Text1 = ""
   Call RECUEC04.Command3_Click
End Sub

Private Sub print2_Click()
   COGEVE04.Option2.Value = True
   Call COGEVE04.Command20_Click
End Sub

Private Sub print3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC04.MESAN = MESASEL$
      AVAFAC04.Option3.Value = True
      AVAFAC04.Option2.Value = True
      Call COGEVE04.Command9_Click
    End If
End Sub

Private Sub print4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC04.MESAN = MESASEL$
      AVAFAC04.Option4.Value = True
      AVAFAC04.Option2.Value = True
      Call COGEVE04.Command9_Click
    End If
End Sub

Private Sub print5_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option2.Value = True
      Call REMEVE04.Command3_Click
    End If
End Sub

Private Sub print6_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option2.Value = True
      Call REMEVE04.Command4_Click
    End If
End Sub

Private Sub print7_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option2.Value = True
      Call REMEVE04.Command5_Click
    End If
End Sub

Private Sub print8_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option2.Value = True
      Call REMEVE04.Command6_Click
    End If
End Sub

Private Sub print9_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option2.Value = True
      Call REMEVE04.Command2_Click
    End If
End Sub

Private Sub recibo_Click()
   Call Command14_Click
End Sub

Private Sub scr1_Click()
  Call COGEVE04.Command7_Click
End Sub

Private Sub scr2_Click()
  COGEVE04.Option1.Value = True
  Call COGEVE04.Command20_Click
End Sub

Private Sub scr3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC04.MESAN = MESASEL$
      AVAFAC04.Option3.Value = True
      AVAFAC04.Option1.Value = True
      Call COGEVE04.Command9_Click
    End If
End Sub

Private Sub scr4_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      AVAFAC04.MESAN = MESASEL$
      AVAFAC04.Option4.Value = True
      AVAFAC04.Option1.Value = True
      Call COGEVE04.Command9_Click
    End If
End Sub

Private Sub scr5_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option1.Value = True
      Call REMEVE04.Command3_Click
    End If
End Sub

Private Sub scr6_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option1.Value = True
      Call REMEVE04.Command4_Click
    End If
End Sub

Private Sub scr7_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option1.Value = True
      Call REMEVE04.Command5_Click
    End If
End Sub

Private Sub scr8_Click()
   REMEVE04.Option1.Value = True
   Call REMEVE04.Command6_Click
End Sub

Private Sub scr9_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      REMEVE04.Option1.Value = True
      Call REMEVE04.Command2_Click
    End If
End Sub

Private Sub sub_iva_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      Call REMEVE04.Command21_Click
    End If
End Sub

Private Sub sub_ven_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      REMEVE04.MESAN = MESASEL$
      Call REMEVE04.Command6_Click
    End If
End Sub

Sub FACABONOS()
    '
5   FACABON04.Text1 = CONTROL$("FACABONO", "LEYTEXTO")
    FACABON04.Text2 = CONTROL$("FACABONO", "IMPORTEM")
    '
10  Screen.MousePointer = 1
    FACABON04.Text3 = "0"
    Call FACABON04.Text3_Change
    FACABON04.Show 1
    If FACABON04.Label2.Caption = "FRESH" Then GoTo 5
    '
    If FACABON04.Label2 <> "OK" Then GoTo 99 ' CANCELADO
    '
    TEXTOFA$ = TRIM$(FACABON04.Text1)
    If TEXTOFA$ = "" Then
         Call MENSERR(24, "Falta Texto Descriptivo de Facturación")
         GoTo 10
       ElseIf Len(TEXTOFA$) > 64 Then
         Call MENSERR(24, "Longitud Excesiva en Texto Descriptivo Factura - Máximo 64 Pos.")
         FACABON04.Text1 = Left$(TEXTOFA$, 64)
         GoTo 10
    End If
    '
    IMPOFA# = XVALO(FACABON04.Text2)
    If IMPOFA# < 0.005 Then
       Call MENSERR(24, "Falta Importe Cuota Mensual")
       GoTo 10
    End If
    '
    Call GRACONTROL("FACABONO", "LEYTEXTO", FACABON04.Text1)
    Call GRACONTROL("FACABONO", "IMPORTEM", FACABON04.Text2)
    '
    For U& = 1 To FACABON04.LICOPEN.ListCount
       If FACABON04.LICOPEN.Selected(U& - 1) = True Then
          NCLIE% = XVALO(Left$(FACABON04.LICOPEN.List(U& - 1), 6))
          If NCLIE% > 0 Then
            If STATCLI%(NCLIE%) > 0 Then
              Call SETULTREG("!CARDEFAC", 0)
              Load FORFACCS
              FORFACCS.BARRELAB.Caption = "BARRE"
              FORFACCS.Text1 = TRIM$(Str$(NCLIE%))
              Call FORFACCS.CARDACLI
              While FORFACCS.Label7.Caption <> "Factura:": Call FORFACCS.Label7_Click: Wend
              While FORFACCS.Label6.Caption <> "(T.Libre)": Call FORFACCS.Label6_Click: Wend
              '
              REBLA$ = REGBLAN$("CARDEFAC")
              Call REPLA(REBLA$, MKS$(1), 1, 4)
              Call REPLA(REBLA$, TEXTOFA$, 7, 64)
              Call REPLA(REBLA$, MKS$(IMPOFA#), 421, 4)
              Call REPLA(REBLA$, MKS$(IMPOFA#), 429, 4)
              Call REPLA(REBLA$, MKD$(IMPOFA#), 433, 8)
              Call GRAREG("!CARDEFAC", REBLA$, 1)
              Call SETULTREG("!CARDEFAC", 1)
              Call CIERRARCH("!CARDEFAC")
              '
              FORFACCS.MARCONOTA.Visible = False
              Call FORFACCS.CARLISFAC
              Call FORFACCS.CALTOFAC
              FORFACCS.Show 1
              If FORFACCS.BARRELAB.Caption = "CANCEL" Then
                 Call COMUNI("Proceso Interrumpido a su Petición.")
                 Unload FORFACCS
                 GoTo 99
              End If
              '
              Call APERBASDAT("CLIEN")
              SQLX$ = "SELECT * FROM Deudor12 "
              SQLX$ = SQLX$ + "WHERE Nume_Cli = " & NCLIE% & " "
              Set ClienTemp = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
              On Error Resume Next
              ClienTemp.MoveFirst
              If Err Then
                    On Error GoTo 0
                 Else
                    On Error GoTo 0
                    ClienTemp.Edit
                    ClienTemp!Mar_Ya_Fac = 1
                    ClienTemp.Update
              End If
              ClienTemp.Close
              '
            End If
          End If
       End If
    Next U&
    '
99  Screen.MousePointer = 1
End Sub
