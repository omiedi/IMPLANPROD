VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COGEVE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Gestión Comercial "
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5220
   FillStyle       =   0  'Solid
   Icon            =   "Cogeve07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   33
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Control Arrastre Saldos"
      Height          =   225
      Left            =   2205
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      FillColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   6945
      Left            =   4320
      ScaleHeight     =   6915
      ScaleWidth      =   1665
      TabIndex        =   25
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   37
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Importar, Descargar y Acceso Directo a Otros Procesos"
         Top             =   3390
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1000
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   310
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Cust."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Acceso a Maestro de Clientes"
         Top             =   2010
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Mast."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "Cogeve07.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   2700
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "VENTAS Y DEUDORES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RESUMENES MENSUALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   240
         TabIndex        =   16
         Top             =   5040
         Width           =   3585
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   630
            ScaleHeight     =   570
            ScaleWidth      =   2295
            TabIndex        =   17
            Top             =   370
            Width           =   2325
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
               Left            =   0
               Picture         =   "Cogeve07.frx":1436
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Generación y Emisión de Listados y Estadísticas Mensuales."
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Subdiarios y Estadística"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   19
               Top             =   80
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONTROL DE ESTADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   3585
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   930
            ScaleHeight     =   570
            ScaleWidth      =   2400
            TabIndex        =   20
            Top             =   1680
            Width           =   2430
            Begin VB.CommandButton Command9 
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
               Left            =   -20
               Picture         =   "Cogeve07.frx":1740
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Listados de Avance de Facturación y Cobranza por Día"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Avance de Factura- cion y Cobranza"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   22
               Top             =   105
               Width           =   1710
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2D3C0&
            Caption         =   "SALIDA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   210
            TabIndex        =   15
            Top             =   2460
            Width           =   3165
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Pantalla"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   36
               Top             =   260
               UseMaskColor    =   -1  'True
               Value           =   -1  'True
               Width           =   1140
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Impresión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1680
               TabIndex        =   35
               Top             =   260
               UseMaskColor    =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   690
            ScaleHeight     =   570
            ScaleWidth      =   2400
            TabIndex        =   12
            Top             =   1260
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
               Picture         =   "Cogeve07.frx":1A4A
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Listado de Saldos y Deudas de Clientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldos y Deudas de Clientes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   14
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   450
            ScaleHeight     =   570
            ScaleWidth      =   2400
            TabIndex        =   9
            Top             =   840
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
               Picture         =   "Cogeve07.frx":1D54
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Listado de Planificación Financiera de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Flujo de Caja Proyectado"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   800
               TabIndex        =   11
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   210
            ScaleHeight     =   570
            ScaleWidth      =   2400
            TabIndex        =   6
            Top             =   420
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
               Picture         =   "Cogeve07.frx":205E
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Listados de Facturas Pendientes de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuentas por Cobrar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   8
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DETALLE POR CUENTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   315
            ScaleHeight     =   570
            ScaleWidth      =   2925
            TabIndex        =   3
            Top             =   370
            Width           =   2955
            Begin VB.CommandButton Command20 
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
               Left            =   2400
               Picture         =   "Cogeve07.frx":2368
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Listado de Límites de Crédito y Deudas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command7 
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
               Picture         =   "Cogeve07.frx":2672
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Cliente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes de Cuenta Corriente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   5
               Top             =   80
               Width           =   1485
            End
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   315
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Cogeve07.frx":297C
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "Cogeve07.frx":2BB0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "COGEVE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
   Command16.Enabled = True
End Sub

Sub Command18_Click()
   Command18.Enabled = False
   If DERACCE%("STPGESVE", "Configuracion de Setup de Ventas") = 2 Then
      'Call CONECRUN("FSETUP04/OPCLAVEN", "Configuración de Funcionamiento")
      OPCLAV07.Show 1
      Call FINAL("")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMASTO", "Archivos Maestros")
    Command19.Enabled = True
End Sub

'
Private Sub Command2_Click()
    CARGAMEN% = 0
    REMEVE07.Show 1
End Sub

Sub Command20_Click()
   Command20.Enabled = False
   Call LILIMCRE
   Command20.Enabled = True
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   Call HELPONLINE("COGEVEN")
   Command29.Enabled = True
End Sub

Sub Command3_Click()
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
'    If ULTREG&("DEUDOR1") > 0 Then
'        X$ = REGLEIDO$("DEUDOR1", 1)      ' fuerza la
'        Call GRAREG("DEUDOR1", X$, 1)     ' reindexacion
'        Call CIERRARCH("DEUDOR1")         ' del archivo
'    End If
    '
    CARFAPEN% = 0
    Call FAPECOB07.FAPECOB1
    '
    FAPECOB07.Show 1
    '
End Sub

Sub Command4_Click()
    '
    Command4.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Or InStr(EPAC$, "AHP") > 0 Then
       FLUCAP07.Command2_Click
       GoTo 99
    End If
    FLUCAP07.Show 1
    '
99  Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALCLI
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   'CONTROL PARA VERIFICAR SALDOS DE CUENTA CORRIENTE VENTAS(POR IMPORTACION DE DATOS (DOSIVAC))
    Exit Sub
    CLAX$ = "": CLACO% = 0
30  PPX% = Int(Rnd * 63)
    If PPX% < 48 Or PPX% > 57 Then GoTo 30
    CLAX$ = CLAX$ + Chr$(PPX%)
    If Len(CLAX$) < 6 Then GoTo 30
    '
    RESPUES$ = CONTRACLA$(CLAX$)
    CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
    If CONTRASE$ <> RESPUES$ Then Exit Sub
    '
    Limit$ = InputBox$("Fecha Limite")
    If Limit$ = "" Then Exit Sub
    FLIMI% = FECHANUM(Limit$)
    If FLIMI% < 1 Then Exit Sub
    '
    FLIM$ = FETEXCOR1$(FLIMI%)
    Dim DEUDA#(32767), SALDO#(32767)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT NuClien, IDEBECOMP, IHABECOMP FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + " and FechEmisi <= '" & FLIM$ & "' "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Set CABATEMP = READSET(SQLX$)
'    Dim CABATEMP As ADODB.Recordset
'    Set CABATEMP = New ADODB.Recordset
'    CABATEMP.CursorLocation = adUseClient
'25  On Error Resume Next
'    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        Exit Sub
      End If
      '
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, FIN&)
        NC% = XVALO(!NUCLIEN)
        If NC% > 0 Then
            IFAC# = XVALO(!idebecomp) - XVALO(!ihabecomp)
            SALDO#(NC%) = SALDO#(NC%) + IFAC#
        End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    FIN& = ULTREG&("FACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIMI% Then
        TI% = Asc(Mid$(X$, 13, 1))
        If TI% < 0 Or TI% > 7 Then TI% = 2
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        IMPO# = CVD(Mid$(X$, 15, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
        IMPO# = SG% * IMPO#
        DEUDA#(NCLIE%) = DEUDA#(NCLIE%) + IMPO#
      End If
    Next U&
    '
    FIN& = ULTREG&("COBRAN")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBRAN", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIMI% Then
      IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
        DEUDA#(NCLIE%) = DEUDA#(NCLIE%) + IMPO#
      End If
    Next U&
    '
    cx& = 32766
    For IL& = 1 To cx&
      Call TRACE(20, IL&, cx&)
      'X$ = REGLEIDO$("DEUDOR1", IL&)
      'NC% = CVI(Left$(X$, 2))
      'If NC% > 0 Then
        ' ajustar saldo de la cuenta
      If Abs(SALDO#(IL&) - DEUDA#(IL&)) >= 0.005 Then
         NC% = IL&
         Y$ = REGBLAN$("LISACLI")
          Call REPLA(Y$, MKI$(NC%), 1, 2)
          Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
          Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
          If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
             Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
          End If
          '
          RELISA& = RELISA& + 1
          Call GRAREG("LISACLI", Y$, RELISA&)
        End If
    Next IL&
    '
    Call SETULTREG("LISACLI", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    Call FINAL("*LISACLI")
    '
    If COMALTER%("Desea Ajustar los Saldos al " + Limit$ + " ? \\Si, Ajustarlos\No, Continuar") <> 1 Then Exit Sub
    '
100 For IL& = 1 To ULTREG&("LISACLI")
        Y$ = REGLEIDO$("LISACLI", IL&)
        NC% = CVI(Left$(Y$, 2))
        FAJUI% = FLIMI%
        SARRASFE# = CVD(Mid$(Y$, 3, 8))
        SALAJU# = CVD(Mid$(Y$, 11, 8))
        CLSS$ = MKI$(NC%) + MKI$(FAJUI%)
        VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(FAJUI%) + "DIFERENCIA DE MIGRACION DE DATOS" + Space$(224) + MKI$(CUEAJU%)
        OBJE$ = "CARAJUCLI"
        '
102     CLSS$ = OBJPLA$(OBJE$, VDEF$)
        If CLSS$ = "" Or CLSS$ = "<ESC>" Then Exit Sub
        '
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
        CUEAJU% = ICONI%
        '
        DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
        DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
        '
        If Abs(DIFESAL#) < 0.005 Then
          If Abs(DIFEDEU#) < 0.005 Then
            Call MENSERR(24, "No hay Ajuste")
            GoTo 98
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
    '    SQLX$ = "SELECT * FROM CAJABANC "
    '    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    '    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    '    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
        
        ORMA% = (-1)
        
        Dim RECITEMP As ADODB.Recordset
        SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
        SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
        SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
        Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        If Not IsNull(RECITEMP!maximo) Then
            ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
        End If
        RECITEMP.Close
        Set RECITEMP = Nothing
    
        'Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
        'With RECITEMP
          'On Error Resume Next
          '.MoveFirst
          'If Err < 1 Then
          '  ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
          'End If
        'End With
        
        ORMX$ = TRIM$(Str$(ORMA% + 1))
        While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
        NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
        '
        MONEMIA% = MONEMI%
        '
20      NC1% = NC%
        Call APERBASDAT("CABAN")
        '
        MONEMI% = MONAJU%
        CoComCor$ = NRECIB$
        CoComLar$ = NRECIB$
'        CodiEmp% = 0
'        CodiSuc% = SUC%
        NumeClie = NC%
        NumeProv% = 0
        FechEmi = CVDAT(FAJUI%)
        '
        Referen$ = rasocli$(NC%)
        FeContab = CVDAT(FAJUI%)
        FeAcredi = CVDAT(FAJUI%)
        NumeTal = 0
        NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
        Observa$ = rasocli$(NC%)
        '
        CuentaCo% = CUECOINT%(CUEMADRE$(NC%))
        TiMovi = "AJUD"            ' AJUSTE DE SALDO PROVEEDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "CCORR"           ' CUENTA CORRIENTE
        Descrip = "Aj.Saldo - " + rasocli$(NC%)
        ImpoDeb# = DIFESAL#
        ValoDola# = ROUND#(DIFESAL# / COTIMONEM#)
        HAYAJU% = 0: If Abs(DIFESAL#) > 0.005 Then HAYAJU% = 1
        Call GRACABAN
        '
    '4.- Graba la Contra-Partida en ICONI%
        CuentaCo% = ICONI%
        TiMovi = "AJUSALD"        ' AJUSTE DE SALDO PROVEEDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "NO-IMP"          ' NO IMPUTADO
        Descrip = "Aj.Saldo - " + rasocli$(NC%)
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
        'CajaBanc.Close
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
        MONEMI% = MONEMIA%
        Screen.MousePointer = 1
        Call COMUNI("Ajuste Completo")
    '
98  Next IL&
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    
End Sub

Sub Command7_Click()
    Command7.Enabled = False
    RECUEC07.Show 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    If EXISTE%("LCOMIS07.EXE") > 0 Then
      OPXX% = ALTERNA%("Liquidación de Comisiones\Otros Accesos Directos")
      If OPXX% = 1 Then
          Call CONECRUN("LCOMIS07", "Liquidación de Comisiones")
          GoTo 99
        ElseIf OPXX% = 2 Then
          GoTo 20
        Else
          GoTo 99
      End If
    End If
    '
20  Call ACCDIR("COGEVEN")
    '
99  Command8.Enabled = True
End Sub

Sub Command9_Click()
    AVAFAC07.Show 1
End Sub

Private Sub Form_Load()
    '
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture6)
    '
    Call ABRECLIEN
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call CIERRARCH("*.*")
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub

Sub COMPACDEU()
   '
   Call CIERRARCH("*.*")
   Call CIERRARCH("!SVD202")
   Call BLANARCH("!SVD202")
   JJ& = 0
   '
   FIN& = ULTREG&("SVD202")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("SVD202", U&)
     IM1# = CVD(Mid$(X$, 11, 8))
     NC% = CVI(Left$(X$, 2))
     If NC% > 0 Then
       If NC% <= 32767 Then
         If Abs(IM1#) >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!SVD202", X$, JJ&)
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("!SVD202")
   '
   On Error GoTo 10
   Kill LUDAT$ + "SVD202.X*"
   Kill LUDAT$ + "SVD202.DAT"
   On Error GoTo 0
   '
   FIN& = ULTREG&("!SVD202")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!SVD202", U&)
     JJ& = JJ& + 1
     Call GRAREG("SVD202", X$, JJ&)
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("*.*")
   Call COMUNI("Compactación Completa.")
   Exit Sub
   '
10 Resume 20
20 On Error GoTo 0
   Call COMUNI("La Compactación no Pudo Completarse.")
   '
End Sub

Sub GELISALCLI()
'
    HOXX = HOY(HO$)
    Dim DEUDA#(999999), SALDO#(999999)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT NuClien, IDEBECOMP, IHABECOMP FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD' "
    '
'    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    '
'    Dim CABATEMP As ADODB.Recordset
'    Set CABATEMP = New ADODB.Recordset
'    CABATEMP.CursorLocation = adUseClient
'25  On Error Resume Next
'    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
    '
    Set CABATEMP = READSET(SQLX$)
    '
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 99
      End If
      '
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, FIN&)
        NCLI = XVALO(!NUCLIEN)
        If NCLI > 0 Then
            IFAC# = XVALO(!idebecomp) - XVALO(!ihabecomp)
            SALDO#(NCLI) = SALDO#(NCLI) + IFAC#
        End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    'jandy sql
    SQLX$ = "SELECT NuClien, salddebe, saldacre FROM SADEUPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Status < 3 "
    '
    'Set SACTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim SACTEMP As ADODB.Recordset
'    Set SACTEMP = New ADODB.Recordset
'    SACTEMP.CursorLocation = adUseClient
'26  On Error Resume Next
'    SACTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 26
'    End If
'    On Error GoTo 0
    '
    Set SACTEMP = READSET(SQLX$)
    With SACTEMP
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        .MoveFirst
        UKL& = 0
        Do While Not .EOF
          UKL& = UKL& + 1
          Call TRACE(20, UKL&, FIN&)
          NCLI = XVALO(!NUCLIEN)
          SALPAR# = XVALO(!SaldDebe) - XVALO(!SALDAcre)
          If Abs(SALPAR#) >= 0.005 Then
            DEUDA#(NCLI) = DEUDA#(NCLI) + SALPAR#
          End If
        .MoveNext
        Loop
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    '
    'cx& = ULTREG&("DEUDOR12")
    RELISA& = 0
    '
    SQLX$ = "select nume_cli from deudor12 with(nolock) where nume_cli>0 order by Raso_Cli asc"
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            NCLI = XVALO(!Nume_Cli)
            ' ajustar saldo de la cuenta
            If Abs(SALDO#(NCLI)) >= 0.005 Or Abs(DEUDA#(NCLI)) >= 0.005 Then
              Y$ = REGBLAN$("LISACLI")
              Call REPLA(Y$, MKS$(NCLI), 1, 4)
              Call REPLA(Y$, MKD$(SALDO#(NCLI)), 5, 8)
              Call REPLA(Y$, MKD$(DEUDA#(NCLI)), 13, 8)
              If Abs(SALDO#(NCLI) - DEUDA#(NCLI)) > 0.005 Then
                 Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 21, 28)
              End If
              '
              RELISA& = RELISA& + 1
              Call GRAREG("LISACLI", Y$, RELISA&)
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
'    For IL& = 1 To cx&
'      Call TRACE(20, IL&, cx&)
'      X$ = REGLEIDO$("DEUDOR1", IL&)
'      NC% = CVI(Left$(X$, 2))
'      If NC% > 0 Then
'        ' ajustar saldo de la cuenta
'        If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
'          Y$ = REGBLAN$("LISACLI")
'          Call REPLA(Y$, MKI$(NC%), 1, 2)
'          Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
'          Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
'          If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
'             Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
'          End If
'          '
'          RELISA& = RELISA& + 1
'          Call GRAREG("LISACLI", Y$, RELISA&)
'        End If
'      End If
'    Next IL&
    '
    Call SETULTREG("LISACLI", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*LISACLI")
       Else
         Call FINAL("*LIMPCLI")
    End If
    '
99 End Sub
'
   '///////*****RUTINA VIEJA******/////
'    '
'    HOXX = HOY(HO$)
'    Dim NC%(16384), DEUDA#(16384), SALDO#(16384)
'    '
'    Screen.MousePointer = 11
'    Call HEADERS("LISACLI", "")
'    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
'    '
'    CX& = ULTREG&("DEUDOR1")
'    '
'    For IL& = 1 To CX&
'      Call TRACE(20, IL&, CX&)
'      X$ = REGLEIDO$("DEUDOR1", IL&)
'      NC%(IL&) = CVI(Left$(X$, 2))
'      SALDO#(IL&) = SALDCLI#(NC%(IL&))
'      DEUDA#(IL&) = 0
'    Next IL&
'    '
'    FIN& = ULTREG&("SVD202")
'    For IL& = 1 To FIN&
'      SS$ = REGLEIDO$("SVD202", IL&)
'      NCI% = CVI(Left$(SS$, 2))
'      TI% = Asc(Mid$(SS$, 3, 1))
'      If TI% > 7 Then TI% = 2
'      SALDEU# = CVD(Mid$(SS$, 11, 8))
'      If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
'      For JL& = 1 To CX&
'        If NC%(JL&) = NCI% Then
'          DEUDA#(JL&) = DEUDA#(JL&) + SALDEU#
'          GoTo 809
'        End If
'      Next JL&
'809 Next IL&
'
'900 RELISA& = 0
'    '
'    For IL& = 1 To CX&
'      If NC%(IL&) > 0 Then
'        If Abs(SALDO#(IL&)) >= 0.005 Or Abs(DEUDA#(IL&)) >= 0.005 Then
'          Y$ = REGBLAN$("LISACLI")
'          Call REPLA(Y$, MKI$(NC%(IL&)), 1, 2)
'          Call REPLA(Y$, MKD$(SALDO#(IL&)), 3, 8)
'          Call REPLA(Y$, MKD$(DEUDA#(IL&)), 11, 8)
'          If Abs(SALDO#(IL&) - DEUDA#(IL&)) > 0.005 Then
'            Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
'          End If
'          '
'          RELISA& = RELISA& + 1
'          Call GRAREG("LISACLI", Y$, RELISA&)
'        End If
'      End If
'    Next IL&
'    '
'    Call SETULTREG("LISACLI", RELISA&)
'    Call CIERRARCH("*.*")
'    Screen.MousePointer = 1
'    '
'    If Option1.Value = True Then
'         Call FINAL("*LISACLI")
'       Else
'         Call FINAL("*LIMPCLI")
'    End If
'    '

Sub LILIMCRE()
    '
    Dim FacPenTemp As ADODB.Recordset
    Dim CABATEMP As ADODB.Recordset
    '
    ' falta considerar deuda en cheques en cartera
    '
    Screen.MousePointer = 11
    Dim DEUCUECO#(999999), CHEPEAC#(999999)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN WITH(NOLOCK) "
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With FacPenTemp
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            NCLIE = !NUCLIEN
            If NCLIE >= 0 Then
              SALDEU# = !SaldDebe - !SALDAcre
              DEUCUECO#(NCLIE) = DEUCUECO#(NCLIE) + SALDEU#
            End If
          .MoveNext
          Loop
       End If
    End With
    '
    ' INCLUYE CHEQUES PENDIENTES DE ACREDITACION
    HOYDIA = FETEXCOR1$(CVINT(Int(Now)))
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND FechAcred >= '" & HOYDIA & "' "
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
    With CABATEMP
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
              If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
              '
              ICHEQ# = !ValAbsComp
              NCLIE = !NUCLIEN
              '
              If NCI% > 0 Then
                If NCI% <= 999999 Then
                  CHEPEAC#(NCLIE) = CHEPEAC#(NCLIE) + ICHEQ#
                End If
              End If
              '
19        .MoveNext
          Loop
       End If
    End With
    '
    FIN& = 0
    Dim rstdeudor As ADODB.Recordset
    Set rstdeudor = FLEXCONN.Execute("select max(nume_cli) as max from deudor12")
    If Not (rstdeudor.EOF And rstdeudor.BOF) Then FIN& = rstdeudor!Max
    rstdeudor.Close: Set rstdeudor = Nothing
    '
    JJ& = 0
    For IL& = 1 To FIN&
       Call TRACE(20, IL&, FIN&)
       NCLIE = IL&    '   CVI(Left$(XX$, 2))
       If Abs(DEUCUECO#(NCLIE)) + Abs(CHEPEAC#(NCLIE)) >= 0.005 Then
          CRELIM# = LIMICRE(NCLIE) 'CVS(Mid$(YY$, 121, 4))
          DISPO# = CRELIM# - DEUCUECO#(NCLIE) - CHEPEAC#(NCLIE)
          EXCES# = 0
          If DISPO# < 0 Then
            EXCES# = Abs(DISPO#)
            DISPO# = 0
          End If
          ZZ$ = REGBLAN$("LILICRE")
          Call REPLA(ZZ$, MKS$(NCLIE), 1, 4)
            RASO$ = rasocli$(NCLIE)
          Call REPLA(ZZ$, RASO$, 5, 30)
          Call REPLA(ZZ$, MKD$(CRELIM#), 35, 8)
          If CRELIM# >= 0 Then
            Call REPLA(ZZ$, MKD$(DEUCUECO#(NCLIE)), 43, 8)
            Call REPLA(ZZ$, MKD$(CHEPEAC#(NCLIE)), 51, 8)
            Call REPLA(ZZ$, MKD$(DISPO#), 59, 8)
            Call REPLA(ZZ$, MKD$(EXCES#), 67, 8)
          End If
          JJ& = JJ& + 1
          Call GRAREG("LILICRE", ZZ$, JJ&)
       End If
    Next IL&
    '
    Call SETULTREG("LILICRE", JJ&)
    Call CIERRARCH("LILICRE")
    '
    If Option1.Value = True Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
    End If
    Screen.MousePointer = 1
    '
End Sub




