VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPORPA07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones Complementarias - Ordenes de Pago"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Left            =   735
      MaskColor       =   &H00404000&
      Picture         =   "OPORPA07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Control B.Datos de Ordenes de Pago"
      Top             =   7770
      Width           =   650
   End
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
      Height          =   640
      Left            =   2310
      Picture         =   "OPORPA07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Analisis de Pagos por Fecha Acreditación de Cheque"
      Top             =   7875
      Width           =   650
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2370
      TabIndex        =   21
      Top             =   7035
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
         Picture         =   "OPORPA07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   75
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   4800
      ScaleHeight     =   6285
      ScaleWidth      =   1260
      TabIndex        =   17
      Top             =   -50
      Width           =   1290
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   41
         Top             =   5160
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   42
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Retenciones de IIBB"
         Top             =   1470
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2985
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command15 
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Análisis Estadístico de Pagos"
         Top             =   4305
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tbls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":0CFC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Acceso Directo a Tabla de Bancos y Valores de Pago"
         Top             =   2320
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":0E46
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   3650
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Height          =   650
         Left            =   105
         MaskColor       =   &H00404000&
         Picture         =   "OPORPA07.frx":1288
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   810
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   650
         Left            =   105
         Picture         =   "OPORPA07.frx":1592
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   150
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "EMISIÓN DE LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   4530
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   14
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":16DC
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Listado de Ordenes de Pago Por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Ordenes de Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   700
            TabIndex        =   16
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   90
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "OPERACIONES COMERCIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4530
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1755
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   31
         Top             =   3045
         Width           =   2430
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   1840
            Picture         =   "OPORPA07.frx":19E6
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Desaplicar Orden de Pago"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":1B30
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Transferencia de Orden de Pago"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Transf. C/C / Desplicar OP"
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
            Left            =   700
            TabIndex        =   32
            Top             =   80
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1485
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   28
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
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":1C7A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Compensación Cliente-Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compensación Cliente-Proveedor"
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
            Left            =   700
            TabIndex        =   29
            Top             =   40
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1200
         ScaleHeight     =   615
         ScaleWidth      =   2400
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
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":1F84
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Ajuste Directo de Saldos Acreedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Saldos Acreedores"
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
            Left            =   700
            TabIndex        =   12
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   40
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   855
         ScaleHeight     =   615
         ScaleWidth      =   2400
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
            Left            =   1840
            Picture         =   "OPORPA07.frx":20CE
            Style           =   1  'Graphical
            TabIndex        =   27
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
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":23D8
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imputación de Créditos y Anticipos a Proveedores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
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
            Left            =   700
            TabIndex        =   9
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   40
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   540
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   1
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command16 
            Caption         =   "Retenciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1350
            TabIndex        =   39
            ToolTipText     =   "Reimpresion de retenciones de Ganancias e IIBB"
            Top             =   0
            Width           =   1050
         End
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
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":2522
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Orden de Pago ya Emitida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver y Imprimir"
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
            Left            =   600
            TabIndex        =   3
            Top             =   30
            Width           =   645
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   300
         ScaleHeight     =   615
         ScaleWidth      =   2400
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
            Height          =   630
            Left            =   0
            Picture         =   "OPORPA07.frx":2B8C
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   700
            TabIndex        =   6
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1530
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPORPA07.frx":2E96
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPORPA07.frx":2F10
      Top             =   1200
   End
End
Attribute VB_Name = "OPORPA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AJUSTEFERVI(PROV%, FF%, IMPORTE#)

    FEX = FF%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    ORMA% = (-1)
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
    End If
    RECITEMP.Close
    Set RECITEMP = Nothing
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
20  NC1% = (-1) * PROV%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    NumeClie = 0
    NumeProv% = PROV%
    FechEmi = CVDAT(FF%)
    '
    Referen$ = rasocli$(NC1%)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = rasocli$(NC1%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1%))
    TiMovi = "AJUP"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = (-1) * IMPORTE#
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALPR"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = IMPORTE#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = IMPORTE#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FF%)
    ImVenc# = IMPORTE#
    Call GRACREPEN

End Sub

Sub ANAPACHEQ()
    '
    Dim INIPER%(12), FINPER%(12)
    '
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    '
    MMXX0$ = TRIM$(MESASEL$)
    MMXX1$ = MESANO$(MMXX0$, DES%, Has%)
    '
    DESA = FETEXCOR1$(DES%)
    HASA = FETEXCOR1$(Has%)
    
    '
    REAVA& = 0
    Screen.MousePointer = 11
    '
    DD1% = DES%
    FEX = DD1%
    SMX0% = SEMANU%(FEX, DELX, ALX)
    INIPER%(0) = DELX: FINPER%(0) = ALX
    For KKI% = 1 To 5
      FEX = DIPOST(ALX)
      SMX1% = SEMANU%(FEX, DELX, ALX)
      INIPER%(KKI%) = DELX: FINPER%(KKI%) = ALX
    Next KKI%
    '
    FEX = INIPER%(3)
    FETE$ = FECHATEX$(FEX)
    MACT% = 1 + Val(Mid$(FETE$, 4, 2))
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(4)
    INIPER%(5) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
10  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 10
    End If
    FINPER%(5) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(5)
    INIPER%(6) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
20  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 20
    End If
    FINPER%(6) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(6)
    INIPER%(7) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
30  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 30
    End If
    FINPER%(7) = 32767
    '
    JJ& = 0

    CTACARVAL% = CUECOINT%(Control$("", "CUECHECA"))
    i& = 0
    SQLX1$ = "(TipoMovim = 'CHENTR' OR TipoMovim= 'RTIB' OR TipoMovim= 'RTGAN') "
    SQLX1$ = SQLX1$ + "AND FECHCONTA >= '" & DESA & "' "
    SQLX1$ = SQLX1$ + " AND FECHCONTA <= '" & HASA & "' "
    FIN& = CantRegistros&("CAJABANC", SQLX1$)
    
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE " & SQLX1$
    SQLX$ = SQLX$ + " ORDER BY FechConta,NUMECOMP ASC"
    '
    Dim CABATMP As ADODB.Recordset
    Set CABATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With CABATMP
      Do While Not .EOF
       Call TRACE(20, i&, FIN&)
       i& = i& + 1
       FECOBRA% = CVINT(!fechconta)
       NCLIE% = XVALO(!NUPROVE)
       NURECI = XVALO(!NumeComp)
       REFERE$ = SAFETEXT$(!DESCRIMOV)
       BANCHE$ = TRIM$(Left$(REFERE$, 18))
       'If BANCHE$ = "" Then BANCHE$ = TRIM$(ECOARCH$("MECOBRA", Left$(X$, 2))) ** COMENTADO P0R QUE NO TENGO EL MEDIO DE COBRANZA PARA PASAR
       NUCHEBA$ = TRIM$(Mid$(REFERE$, 19, 14))
       If NUCHEBA$ <> "" And TRIM$(Mid$(REFERE$, 19, 14)) <> "" Then NUCHEBA$ = NUCHEBA$ + " - "
       NUCHEBA$ = NUCHEBA$ + BANCHE$
       'If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
       IMPOCHE# = XVALO(!IHABECOMP)
       FACRE% = CVINT(!fechacred)
       
       If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
       
 ' ESTO QUE SIGUE QUEDA IGUAL
       Y$ = REBLA$
       Call REPLA(Y$, MKS$(NURECI), 1, 4)
       Call REPLA(Y$, MKI$(FECOBRA%), 5, 2)
       Call REPLA(Y$, MKI$(NCLIE%), 7, 2)
       Call REPLA(Y$, NUCHEBA$, 9, 32)
       Call REPLA(Y$, MKI$(FACRE%), 41, 2)
       For KKI% = 0 To 7
            If FACRE% <= FINPER%(KKI%) Then
              Call REPLA(Y$, MKD$(IMPOCHE#), 43 + 8 * KKI%, 8)
              GoTo 40
            End If
       Next KKI%
40     Call REPLA(Y$, MKD$(IMPOCHE#), 107, 8)
       JJ& = JJ& + 1
       Call GRAREG("ANACHEMI", Y$, JJ&)
      .MoveNext
     Loop
    End With

    '
    Call SETULTREG("ANACHEMI", JJ&)
99  Call CIERRARCH("*.*")
    Call HEADERS("ANACHEMI", "")
' EL MMXX1 ES EL PERIO$
    Call HEADERS("ANACHEMI", "Corresponde a: " + MMXX1$)
' DEBERIA FUNCIONAR

    Call FINAL("*LIACHEMI")
    Screen.MousePointer = 1
   '
End Sub


Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

 Sub Command1_Click()
    Command1.Enabled = False
    Call STORPA07.Command13_Click
    '
390 Command1.Enabled = True
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
    'Se agregó listado por fechas(opx=5)
    OPX% = COMALTER%("Puede Visualizar Aplicaciones de Diversas Formas:\\Por Orden de Pago o N/C\Por Factura o N/D\Listado por Rango de Fechas")
    If OPX% = 1 Then
        Call VAPLIOP
      ElseIf OPX% = 2 Then
        Call VAPLIFAC
      ElseIf OPX% = 3 Then
         Call LIAPLIFE
    End If
    Command11.Enabled = True
End Sub

Sub Command12_Click()
    Command12.Enabled = False
    If DERACCE%("DESAPOPA", "Desaplicar Ordenes de pago") > 1 Then
      Call DESAPLIOPA         ' Des-aplicacion de Orden de Pago
    End If
    Command12.Enabled = True
End Sub

Sub Command13_Click()
    Command13.Enabled = False
    If DERACCE%("TRANORPA", "Transferencia de Ordenes de pago") > 1 Then
      Call TRANSORPA             ' TRANSFERENCIA DE ORDENES DE PAGO
    End If
    Command13.Enabled = True
End Sub
Private Sub Command14_Click()
   Command14.Enabled = False
    If DERACCE%("STPORPAG", "Configuración de Funcionamiento de Ordenes de Pago") = 2 Then
        STORPA07.Show 1
        Call COMUNI("Re-Inicie Ahora la Aplicación.")
        Call FINAL("")
    End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   '
   Call ANAPACHEQ
   '
   Command15.Enabled = True
   
End Sub

Sub Command16_Click()

     Command16.Enabled = False
     '
2    OPX% = COMALTER%("Elija el Tipo de Retención\\Ingresos Brutos\Ganancias\I.V.A.")
     If OPX% < 1 Or OPX% > 3 Then GoTo 99
     
     If OPX% = 1 Then
        OpcionesIIBB% = COMALTER%("Jurisdicciones:\\Buenos Aires\C.A.B.A.\Misiones\Tucumán")
        If OpcionesIIBB% < 1 Or OpcionesIIBB% > 4 Then GoTo 2
        If OpcionesIIBB% = 1 Then
            Call CARGALISEL("LIIBBSAS", "CAJABANC", "NUMECOMP/I9;DESCRIMOV/A32", "Tipomovim='RTIB' AND OPCIMOVIM='Bs-As' order by numecomp desc")
10          Call SELECHO("LIIBBSAS/Retenciones de IIBB Bs As")
            '
            IDRET& = XVALO(Mid$(VALACT2$("LIIBBSAS"), 14))
            If IDRET& > 0 Then
                Screen.MousePointer = 11
                NURETBA& = IDRET&
                Condi$ = "descrimov like '%" & NURETBA& & "%' and tipomovim='RTIB'"
                OPA$ = VALOBADA("cajabanc", "codicom_lar", Condi$)
                RETEIIBB# = XVALO(VALOBADA("cajabanc", "valabscomp", "codicom_lar='" & OPA$ & "' and tipomovim='RTIB'"))
                FF% = CVINT%(VALOBADA("cajabanc", "fechconta", "codicom_lar='" & OPA$ & "' and tipomovim='OPAG'"))
                Base# = XVALO(VALOBADA("cajabanc", "valabscomp", "codicom_lar='" & OPA$ & "' and tipomovim='OPAG'"))
                NPROVEI% = VALOBADA("cajabanc", "nuprove", "codicom_lar='" & OPA$ & "' and tipomovim='OPAG'")
                '
                FECHDOC$ = FETEXTO$(FF%)
                NUMEDOC$ = TRIM$(Str$(NURETBA&))
                While Len(NUMEDOC$) < 8
                   NUMEDOC$ = "0" + NUMEDOC$
                Wend
                NUMEDOC$ = "0001-" + NUMEDOC$
                TIPODOC$ = "Retencion I.Brutos"
                '
                CODPARAM$(1) = "IMPO-CHE"
                CODPARAM$(2) = "IMPO-PEN"
                CODPARAM$(3) = "IMPO-CAN"
                CODPARAM$(4) = "POR-DESC"
                CODPARAM$(5) = "DOCU-APL"
                CODPARAM$(6) = "TIPO-VAL"
                CAPARDOC% = 6
                '\\\OT-05-0514-WAR-14/7/05///
                  ' determinar categoria de I.Brutos
                CATEIB2$ = ""
                If TICUEPRO%((-1) * NPROVEI%) = 0 Then ' RETENCION CERO EN CUENTAS DE GASTOS
                  RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPROVEI%))
                  CATIB1% = Asc(Mid$(RERET$, 3, 1))
                  CATEIB2$ = ECOARCH("TACATIBP", Chr$(CATIB1%))
                  PORCONVE = CVS(Mid$(RERET$, 5, 4))
                  PORCEIBRU# = ALIRETBRU#(NPROVEI%, MINOIMPO#, MEDIPAGO%)
                End If
                '\\\OT-05-0514-WAR-FIN///
                DOCPARAM$(1) = TRIM$(FORMATNUM$(Base#, "F12.2"))
                DOCPARAM$(2) = TRIM$(FORMATNUM$(PORCEIBRU#, "F12.2"))
                DOCPARAM$(3) = TRIM$(FORMATNUM$(RETEIIBB#, "F12.2"))
                DOCPARAM$(4) = TRIM$(FORMATNUM$(PORCONVE, "F12.2"))
                DOCPARAM$(5) = TRIM(Text11)
                DOCPARAM$(6) = TRIM(CATEIB2$) 'CATEGORIA
                '
                CACOLTAB1% = 0
                '
                DESTIDOCNUE = (-1) * NPROVEI%
                CAITAB1% = 0
                '
                CANCELPRINT% = 1 'PARA QUE LO VISUALICE POR PANTALLA Y DECIDA SI LO IMPRIME Y PUEDA CAMBIAR DE IMPRESORA
                Call PRINTDOC("FORRETIB")   ' Retencion I.Brutos
                DOCUNU& = 0
                Call MUESTRADOC(DOCUNU&, "", TIPODOC$, TRIM$(Str$(NURETBA&)), NUMEDOC$)
                DOCUNU& = 0: TIPODOC$ = "": NUMEDOC$ = 0
                CANCELPRINT% = 0
                '
            End If
            
        ElseIf OpcionesIIBB% = 2 Then
            Call CARGALISEL("LIIBCABA", "CAJABANC", "NRODERETENCION/I9;OBSERGEN/A64", "Tipomovim='RTIB' AND OPCIMOVIM='CABA' GROUP BY NRODERETENCION, OBSERGEN order by NRODERETENCION")
            Call SELECHO("LIIBCABA/Retenciones de IIBB CABA")
            
            NroRetencion$ = VALACT1$("LIIBCABA")
            If XVALO(NroRetencion$) > 0 Then
                FormularioIIBB$ = TRIM$(Control$("", "FOREIBCA"))
                
                CondicionSql$ = "NRODERETENCION=" & NroRetencion$ & " AND Tipomovim='RTIB' AND OPCIMOVIM='CABA'"
                While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
                Proveedor& = XVALO(VALOBADA("CAJABANC", "NUPROVE", CondicionSql$))
                FechaRetencion% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", CondicionSql$))
                
                FILTX$ = NroRetencion$ & ";" & CStr(Proveedor&) & ";" & FECHATEX$(FechaRetencion%) & ";" & CStr(XVALO(NroRetencion$))
                Call STDFORM(FormularioIIBB$, FILTX$)
            End If
            
        ElseIf OpcionesIIBB% = 3 Then
            Call CARGALISEL("LIIBMISI", "CAJABANC", "NRODERETENCION/I9;OBSERGEN/A64", "Tipomovim='RTIB' AND OPCIMOVIM='MISIONES' GROUP BY NRODERETENCION, OBSERGEN order by NRODERETENCION")
            Call SELECHO("LIIBMISI/Retenciones de IIBB MISIONES")
            
            NroRetencion$ = VALACT1$("LIIBMISI")
            If XVALO(NroRetencion$) > 0 Then
                FormularioIIBB$ = TRIM$(Control$("", "FOREIBMI"))
                
                CondicionSql$ = "NRODERETENCION=" & NroRetencion$ & " AND Tipomovim='RTIB' AND OPCIMOVIM='MISIONES'"
                While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
                Proveedor& = XVALO(VALOBADA("CAJABANC", "NUPROVE", CondicionSql$))
                FechaRetencion% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", CondicionSql$))
                
                FILTX$ = NroRetencion$ & ";" & CStr(Proveedor&) & ";" & FECHATEX$(FechaRetencion%) & ";" & CStr(XVALO(NroRetencion$))
                Call STDFORM(FormularioIIBB$, FILTX$)
            End If
            
         Else
            Call CARGALISEL("LIIBMISI", "CAJABANC", "NRODERETENCION/I9;OBSERGEN/A64", "Tipomovim='RTIB' AND OPCIMOVIM='TUCUMAN' GROUP BY NRODERETENCION, OBSERGEN order by NRODERETENCION")
            Call SELECHO("LIIBMISI/Retenciones de IIBB TUCUMAN")
            
            NroRetencion$ = VALACT1$("LIIBMISI")
            If XVALO(NroRetencion$) > 0 Then
                FormularioIIBB$ = TRIM$(Control$("", "FOREIBTU"))
                
                CondicionSql$ = "NRODERETENCION=" & NroRetencion$ & " AND Tipomovim='RTIB' AND OPCIMOVIM='TUCUMAN'"
                While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
                Proveedor& = XVALO(VALOBADA("CAJABANC", "NUPROVE", CondicionSql$))
                FechaRetencion% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", CondicionSql$))
                
                FILTX$ = NroRetencion$ & ";" & CStr(Proveedor&) & ";" & FECHATEX$(FechaRetencion%) & ";" & CStr(XVALO(NroRetencion$)) & ";ORIGINAL"
                Call STDFORM(FormularioIIBB$, FILTX$)
            End If
         
        End If
        
        GoTo 2
        
    ElseIf OPX% = 2 Then
        Call CARGALISEL("LISGANA", "CAJABANC", "NRODERETENCION/I9;OBSERGEN/A64", "Tipomovim='RTGAN' GROUP BY NRODERETENCION, OBSERGEN order by NRODERETENCION")
20      Call SELECHO("LISGANA/Retenciones Ganancias")
        '
        IDRET& = XVALO(VALACT1$("LISGANA"))
        If IDRET& > 0 Then
            NroRetencion$ = CStr(IDRET&)
            While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
            DOCUNU& = DOCID&("Retencion Ganancias", CStr(IDRET&), NroRetencion$)
            Call MUESTRADOC(DOCUNU&)
        End If
        
   ElseIf OPX% = 3 Then
   
        FormularioIVA$ = TRIM$(Control$("", "FORETIVA"))
        If FormularioIVA$ = "" Then
            Call MENSERR(24, "Imposible Imprimir Comprobante.\ \Falta Configurar Formulario de Impresiíon")
            GoTo 2
        End If
                
        Call CARGALISEL("LIRETIVA", "CAJABANC", "NRODERETENCION/I9;OBSERGEN/A38;CODICOM_LAR/A18", "Tipomovim='RTIVA' GROUP BY NRODERETENCION, OBSERGEN, CODICOM_LAR order by NRODERETENCION")
30      Call SELECHO("LIRETIVA/Retenciones de IVA")
        
        NroRetencion$ = VALACT1$("LIRETIVA")
        If XVALO(NroRetencion$) > 0 Then
            OrdenPago$ = TRIM$(Mid$(VALACT2$("LIRETIVA"), 40))
            CondicionSql$ = "TIPOMOVIM='RTIVA' AND CODICOM_LAR='" & OrdenPago$ & "' AND NRODERETENCION=" & NroRetencion$
            
            FechaRetencion% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", CondicionSql$))
            ANIO$ = Format(CVDAT(FechaRetencion%), "yyyy")
            
            While Len(NroRetencion$) < 8: NroRetencion$ = "0" & NroRetencion$: Wend
            NroRetencion$ = "0000-" & ANIO$ & "-" & NroRetencion$
            
            Proveedor& = XVALO(VALOBADA("CAJABANC", "NUPROVE", CondicionSql$))
            
            FILTX$ = NroRetencion$ & ";" & CStr(Proveedor&) & ";" & FECHATEX$(FechaRetencion%) & ";" & CStr(XVALO(Mid$(NroRetencion$, 11)))
            Call STDFORM(FormularioIVA$, FILTX$)
                      
            GoTo 30
        End If
        GoTo 2
   End If
   '
99 Command16.Enabled = True
    
End Sub


Sub Command17_Click()
     Command17.Enabled = False
   
     Call CARGALISEL("PRINTRIB", "PDOC_TBL", "NUMEDOC/I10;REFERENCIA/A64", "TIPODOCU = 'Retencion I.Brutos' AND REFERENCIA LIKE '%Retencion I.Brutos%'", "DISTINCT")
     '
5    Call SELECHO("PRINTRIB/Retenciones de IIBB")
     RESUL& = XVALO(VALACT1$("PRINTRIB"))
     If RESUL& > 0 Then
        DOCUNU& = XVALO(VALOBADA("PDOC_TBL", "NUMEDOC", "NUMEDOC='" & RESUL& & "' ORDER BY NUMEDOC DESC"))
        Call MUESTRADOC(DOCUNU&)
        GoTo 5
     End If
     '
     Command17.Enabled = True
     Exit Sub
   
''   If CLACONTRA% < 1 Then                      ' iNGRESO DE CLAVE Y CONTRASEÑA
''       Call MENSERR(24, "Santo Y Seña Inválido")
''       Exit Sub
''   End If
'
'10 AA$ = InputBox("Ingrese Orden de Pago", "Anulacion Retencion IIBB")
'   AB& = XVALO(AA$)
'   If AB& = 0 Then GoTo 99
'
'   For U& = 1 To ULTREG("SURETIB")
'        Z1$ = REGLEIDO$("SURETIB", U&)
'        nuorpax& = CVS(Mid$(Z1$, 7, 4)) 'LEE LA ORDEN DE PAGO
'        If nuorpax& = AB& Then
'            OPX% = COMALTER%("Desea Eliminar la Retencion de la O/P '" & nuorpax& & "'\\No, Cancelar\Si, Continuar")
'            If OPX% = 2 Then
'                Call REPLA(x$, String$(16, 0), 79, 16)
'                Call GRAREG("SURETIB", x$, U&)
'                Exit For
'            End If
'        End If
'    Next U&
'
'    Call COMUNI("Proceso Finalizado")
'    GoTo 10
99  Command17.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Hide
    '
End Sub
'
Sub Command23_Click()
    '
    ' MODIFICADO 9/4/11 POR EPB PARA READSET Y ELIMINA DUPLICADOS
    Command23.Enabled = False
    
2   OPX% = COMALTER%("Opciones de Impresión\\Indice de Ordenes de Pago\Por Número de Orden de Pago")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    If OPX% = 1 Then
        '
        Call ORPAGO07.ACTULISTAOP
        ' LA RUTINA DE CARGA DE LA LISTA DE ORDENES DE PAGO SE LLEVO AL FRM ORPAGO07 RUTINA ACTULISTAOP
        Call TRALOCAL("INDIORPA", "")
        '
3       Call SELECHO("!INDIORPA/Indice de Ordenes de Pago Emitidas")
        NPX& = XVALO(VALACT1$("!INDIORPA"))
        If NPX& < 1 Then GoTo 2
        '
        Condi$ = "TipoMovim = 'OPAG' AND CODICOM_LAR LIKE 'OP%' AND NUMECOMP= " & NPX&
        Stat% = VALOBADA("CAJABANC", "STATUS", Condi$, "NOMESS")
        If Stat% = 9 Then Call COMUNI("La Orden de Pago " & NPX& & " se Encuentra Anulada")
        '
        TIDOCUM$ = "Orden de Pago"
        NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
        DOCUNU& = 0
5       Call MUESTRADOC(DOCUNU&, , TIDOCUM$, NUDOCU$, NUDOCU$)
10      GoTo 3

    ElseIf OPX% = 2 Then
        '
13      NPX& = XVALO(InputBox("Ingrese Número de Orden de Pago", "Orden de Pago", ""))
        If NPX& < 1 Then GoTo 2
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
              GoTo 15
            End If
          End If
         Next UI&
         Screen.MousePointer = 1
         Call COMUNI("Documento no Encontrado")
         GoTo 13
15       Call MUESTRADOC(DOCUNU&)
         GoTo 13
         Screen.MousePointer = 1
    End If
    '
99  Command23.Enabled = True
    '
End Sub

 Sub Command25_Click()
    Command25.Enabled = False
    Call ANUORPA         ' Anulación de Orden de Pago
    Command25.Enabled = True
End Sub

 Sub Command3_Click()
    Command3.Enabled = False
    Call FACOCOMP07.Command3_Click
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    Call AJUSALP
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
    
    Dim CAMP%, CUEMP%(2048), MEPA%(2048)
    Dim NOPFF(2048)
    '
    Command5.Enabled = False
    '
    'Agregue
    NUEFORMATO% = 0
    If EXISTE%(LUDAT$ + "LISOPAG.RFS") > 0 Then
        NUEFORMATO% = 1
        REBLB$ = REGBLAN$("LISOPAG")
    End If
    '\\\OT-08-0875-OD-19/08/05///
    OPXX% = ALTERNA%("O/P's por Rango de Fechas\ Anulaciones por Rango De Fechas")
    If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
    If OPXX% = 1 Then
      GoTo 10
     Else
      Call FINAL("?OPAGANULADA")
      GoTo 99
    End If
    '\\\OT-08-0875-OD-FIN///
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If DES% > Has% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    DESDAT = FETEXCOR1$(DES%)
    HASDAT = FETEXCOR1$(Has%)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
    End If
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "Order by FechConta, substring(CodiCom_Lar,4,11) "
    SQLX$ = SQLX$ + ", NumeComp  "
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    J1& = 0: J2& = 0: REBLA$ = REGBLAN$("DETPAGC")
    NIVCONXXX% = NIVCONFI%
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      '
      Do While Not .EOF
        '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
        If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 19
        '\\\OT-08-1032-FL-FIN
        If !CODICOM_LAR = NUMEFA$ Then
          If !NUPROVE = NCLIE% Then
            If Int(CDbl(!fechconta)) = FECHAD Then
              INETO# = INETO# + !Idebecomp - !IHABECOMP
              ITOTA# = ITOTA# + !Idebecomp - !IHABECOMP
              Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
              Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
              Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
              Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
              Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
              Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
              Call GRAREG("DETPAGC", ZZ$, J1&)
              '
              If NUEFORMATO% = 1 Then
                TTA$ = REBLB$
                TT1$ = Left$(ZZ$, 32)
                Call REPLA(TTA$, TT1$, 1, 32)
                Call REPLA(TTA$, MKD$(ITOTA#), 65, 8)
                Call REPLA(TTA$, TRIM(!OBSERGEN), 73, 48)
                Call GRAREG("LISOPAG", TTA$, J1&)
              End If
              '
              GoTo 19
            End If
          End If
        End If
        '
        ZZ$ = REBLA$
        NUMEFA$ = !CODICOM_LAR
        NCLIE% = !NUPROVE
        FECHAD = Int(CDbl(!fechconta))
        INETO# = !Idebecomp - !IHABECOMP ' !ValAbsComp
        ITOTA# = !Idebecomp - !IHABECOMP ' !ValAbsComp
        '
        Call REPLA(ZZ$, Left$(TRIM$(!TipoMovim), 2), 1, 2)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, !CODICOM_LAR, 5, 18)
        Call REPLA(ZZ$, MKI$(!NUPROVE), 23, 2)
        Call REPLA(ZZ$, MKD$(INETO#), 25, 8)
        Call REPLA(ZZ$, MKD$(IIVA#), 33, 8)
        Call REPLA(ZZ$, MKD$(IVADU#), 41, 8)
        Call REPLA(ZZ$, MKD$(RETEN#), 49, 8)
        Call REPLA(ZZ$, MKD$(IOTRO#), 57, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 65, 8)
        '
        J1& = J1& + 1
        Call GRAREG("DETPAGC", ZZ$, J1&)
        '
        'AGREGUE ACA
        If NUEFORMATO% = 1 Then
          TTA$ = REBLB$
          TT1$ = Left$(ZZ$, 32)
          Call REPLA(TTA$, TT1$, 1, 32)
          Call REPLA(TTA$, MKD$(ITOTA#), 65, 8)
          Call REPLA(TTA$, TRIM(!OBSERGEN), 73, 48)
          Call GRAREG("LISOPAG", TTA$, J1&)
        End If
19    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    Call SETULTREG("DETPAGC", J1&)
    Call HEADERS("DETPAGC", "")
    Call HEADERS("DETPAGC", "Corresponde a: " + TRIM$(PERIO$))
    
    '
    If NUEFORMATO% = 1 Then
       Call SETULTREG("LISOPAG", J1&)
       Call HEADERS("LISOPAG", "")
       Call HEADERS("LISOPAG", "Corresponde a: " + TRIM$(PERIO$))
       RTA1% = COMALTER%("Seleccione la Formato de Listado:\\Tradicional\Con Referencia")
       If RTA1% = 2 Then
         Call FINAL("*LISOPAG")
         GoTo 99
       End If
    End If
    Call FINAL("*LIDPAGC")
    '
99  Screen.MousePointer = 1
    Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call DESAPLIOPA         ' Des-aplicacion de Orden de Pago
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    'Call HELPONLINE("ORPASTA")
    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Call AJUSTEFERVI(583, FECHANUM("31/12/12"), -0.05)
        Call AJUSTEFERVI(832, FECHANUM("31/12/12"), -48.5)
        Call AJUSTEFERVI(976, FECHANUM("31/12/12"), -0.01)
        Call AJUSTEFERVI(1343, FECHANUM("28/12/12"), 55.12)
        Call AJUSTEFERVI(1537, FECHANUM("20/04/12"), 2606.39)
        Call AJUSTEFERVI(1671, FECHANUM("20/12/12"), 0.02)
        
        Call COMUNI("Proceso Finalizado")
    End If
End Sub

Sub DESAPLIOPA()       ' DES-APLICACION de ordenes de pago
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
     '
     FECHALIM% = XVALO(Control$("", "CIEMECOM"))
     '
10   Call SELECHO("ACREDOR")
     NC% = XVALO(VALACT1$("ACREDOR"))
     If NC% < 1 Then Exit Sub
     '
100  Call SETULTREG("!SELCUECO", 0)
     JJ& = 0
     ' REEMPLAZAR POR BUSQUEDA SQL EN EL CAJABANC QUE OBTENGA TODAS LAS ORDENES
     ' DE PAGO Y ARME EL SELCUECO. IGUAL AL DE TRANSFERENCIA
     FechaLimite = FETEXCOR1$(FECHALIM%)
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ & " WHERE (TipoMovim = 'OPAG' OR TIPOMOVIM = 'AJUP' or (tipomovim='fcom' and opcimovim='nc'))"
     SQLX$ = SQLX$ & " AND NuProve  = " & NC% & ""
     SQLX$ = SQLX$ & " AND FechEmisi  > '" & FechaLimite & "'"
'     SQLX$ = SQLX$ & " ORDER BY  cast(substring(CodiCom_Cor,4,6) as int) "
     SQLX$ = SQLX$ & " ORDER BY  NUMECOMP ASC "
     Dim CAJABTMP As ADODB.Recordset
     Set CAJABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CAJABTMP
       Do While Not .EOF
        Call REPLA(x$, !CODICOM_LAR, 7, 18)
        IMPO# = !Idebecomp - IHABECOMP
        If IMPO# >= 0.005 Then
           JJ& = JJ& + 1
           IMPOS$ = FORMATNUM$(IMPO#, "F12.2")
           Y$ = MKS$(JJ&) + Mid$(x$, 7, 18) + IMPOS$
           Call GRAREG("!SELCUECO", Y$, JJ&)
        End If
        .MoveNext
       Loop
     End With
     CAJABTMP.Close
     Set CAJABTMP = Nothing
     Call SETULTREG("!SELCUECO", JJ&)
     Screen.MousePointer = 1
     '
     Call CIERRARCH("!SELCUECO")
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Ordenes de Pago que Puedan Desaplicarse.")
        GoTo 10
     End If
     '
     Call FRESHECHO("!SELCUECO")
     '
     Call SELECHO("!SELCUECO")
     If VALACT1$("!SELCUECO") = "" Then GoTo 10
     DOCUNU$ = TRIM(Mid$(VALACT2$("!SELCUECO"), 1, 18))
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ & " WHERE TipoMovim = 'OPAG' OR TIPOMOVIM = 'AJUP'"
     SQLX$ = SQLX$ & " AND NuProve  = " & NC% & ""
     SQLX$ = SQLX$ & " AND CodiCom_Lar  = '" & DOCUNU$ & "'"
     
     Dim CAJABTMP1 As ADODB.Recordset
     Set CAJABTMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CAJABTMP1
        IMPO# = XVALO(!Idebecomp - IHABECOMP)
        FEDOC% = CVINT(!FECHEMISI)
        NC% = XVALO(!NUPROVE)
     End With
     CAJABTMP1.Close
     Set CAJABTMP1 = Nothing
     '
     Screen.MousePointer = 1
     IMPOS$ = FORMATNUM$(IMPO#, "F12.2") 'Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     '
     'AQUI Se Debe Intervenir
     SQL$ = "SELECT * FROM SACREPEN WHERE NUPROVE = 0"
     Dim SACRE As ADODB.Recordset
     Set SACRE = New ADODB.Recordset
     SACRE.Open SQL$, FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText
     '
     SQL$ = "SELECT * FROM APLIPAGO with(NOLOCK)"
     SQL$ = SQL$ + " WHERE CoDocAp_Lar = '" & DOCUNU$ & "'"
     Dim APLIP As ADODB.Recordset
     Set APLIP = New ADODB.Recordset
     APLIP.Open FILTSQL$(SQL$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With APLIP
        If .EOF = True And .BOF = True Then
           Call COMUNI("No se ha Encontrado Aplicaciones\Relacionadas con el Comprobante " & DOCUNU$)
           GoTo 10
        End If
        Do While Not (.EOF)
            'OBENTER DATOS DE APLIPAGO y
            Condi$ = "codicom_lar = '" & DOCUNU$ & "'"
            FECHA$ = VALOBADA("CAJABANC", "FECHCONTA", Condi$)
            'AGREGAR REGISTROS EN SACREPEN
            SACRE.AddNew
            SACRE!NUCLIEN = 0
            SACRE!NUPROVE = XVALO(!NUPROVE)
            SACRE!RaSoCliPro = SAFETEXT(!RASOPROV)
            SACRE!FECHEMISI = FECHA$
            SACRE!CODICOM_LAR = SAFETEXT(!codocor_lar)
            SACRE!CODICOM_COR = SAFETEXT(!CODOCOR_COR)
            SACRE!IBruComp = XVALO(!IMPOAPLI)
            SACRE!IIvaComp = XVALO(!IIVAORIG)
            SACRE!FEVENCIM = FECHA$
            SACRE!SaldDebe = 0
            SACRE!SaldAcre = XVALO(!IMPOAPLI)
            SACRE!Observa = "Desaplicacion Manual - " & USERNAME$ & " - " & Date & " " & Time
            SACRE!CodiEmpr = CodiEmpr%
            SACRE!CodiSucu = 0
            SACRE!Niv_Clasi = NiveCla%   ' de donde se saca esto ?
            '
            Condi$ = "codicom_lar = '" & DOCUNU$ & "'"
            FECHA$ = VALOBADA("CAJABANC", "FECHCONTA", Condi$)
            'AGREGAR REGISTROS EN SACREPEN
            SACRE.AddNew
            SACRE!NUCLIEN = 0
            SACRE!NUPROVE = XVALO(!NUPROVE)
            SACRE!RaSoCliPro = SAFETEXT(!RASOPROV)
            SACRE!FECHEMISI = FECHA$
            SACRE!CODICOM_LAR = DOCUNU$
            SACRE!CODICOM_COR = SAFETEXT(!CODOCAP_COR)
            SACRE!IBruComp = XVALO(!IMPOAPLI)
            SACRE!IIvaComp = 0
            SACRE!FEVENCIM = FECHA$
            SACRE!SaldDebe = XVALO(!IMPOAPLI)
            SACRE!SaldAcre = 0
            SACRE!CodiEmpr = CodiEmpr%
            SACRE!CodiSucu = 0
            SACRE!Niv_Clasi = NiveCla%    ' de donde se saca esto ?
            SACRE!Observa = "Desaplicacion Manual - " & USERNAME$ & " - " & Date & " " & Time
            SACRE.Update
        .MoveNext
        Loop
     End With
     SACRE.Close
     Set SACRE = Nothing
     APLIP.Close
     Set APLIP = Nothing
     
     Condi$ = "CoDocAp_Lar = '" & DOCUNU$ & "'"
     Call BORRAREGISTROS("APLIPAGO", Condi$, REG&)
     Call COMUNI("Desaplicacion Finalizada Correctamente")
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
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If DES% > Has% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    JJ& = 0
    '
    If CAMP% <= 0 Then
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
    End If
    '
    CANOPFF% = 0
    CANOPAI% = 0
    FIN& = ULTREG&("CUECORR")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", i&)
      If CVI(Left$(CC0$, 2)) >= DES% Then
        If CVI(Left$(CC0$, 2)) <= Has% Then
          If UCase$(Mid$(CC0$, 12, 4)) <> "ANUL" Then
            NOPA& = XVALO(Mid$(CC0$, 18, 8))
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
   Next i&
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
        If FEOPA% <= Has% Then
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
            VUELTA% = 0
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
            If VUELTA% = 0 Then
              VUELTA% = 1
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
       Call MENSERR(24, "Error en Orden de Pago" + Str$(NOPA&) + "\de " + TRIM$(rasocli$((-1) * NP%)) + "\por " + TRIM$(CSTRING$(MKD$(IOPA#(KKI%)), 4, 12, 2, 2)) + " del " + FECHATEX$(FEX))
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
Call ANUORPA2
Exit Sub
    If DERACCE%("ANUORPA", "Anulación de Ordenes de Pago") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'OPAG' OR (TipoMovim = 'AJUP' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + "AND NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND INT(FechEmisi) > " & FECHALU & " "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    NIVCONXXX% = NIVCONFI%
    
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          '\\\OT-08-1032-FL-12/11/08 VERIFICA EL NIVEL DEL USUARIO PARA
          If XVALO(!Niv_Clasi) > NIVCONXXX% Then GoTo 70
          '\\\OT-08-1032-FL-FIN
          NCLIE% = !NUPROVE
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then ' PARA QUE TAMBIEN TRAIGA LAS ANULADAS
          If UCase$(Left$(!CODICOM_LAR, 2)) = "OP" Or UCase$(Left$(!CODICOM_LAR, 2)) = "CC" Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            Y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
70      .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Pago que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
12  Call SELECHO("!SELCUECO")
    If TRIM$(VALACT2$("!SELCUECO")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUECO"), 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6)) ' para anular retenciones i.brutos
    '
    MODOANU% = COMALTER%("Modo de Anulación del Comprobante '" + DOCULA$ + "'\\Marca de Anulación\Re-Utilizar Numero")
    If MODOANU% < 1 Or MODOANU% > 2 Then GoTo 12
    '
    CCC$ = ""
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    With CajaBanc
19    .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
20    If .NoMatch = False Then
        ' DATOS PARA LA RECONSTRUCCION
'        CodiEmp% = !CodiEmpr
'        CodiSuc% = !CODISUCU
        '
        If MODOANU% = 2 Then
           CCC$ = Left$(TRIM$(!CODICOM_COR), 9) + "-nl"
          .Delete
          GoTo 19
        End If
        '
        .Edit
           CCL$ = !CODICOM_LAR
           Call REPLA(CCL$, "anul", 6, 4)
        !CODICOM_LAR = CCL$
           CCC$ = Left$(TRIM$(!CODICOM_COR), 9) + "-nl"
        !CODICOM_COR = CCC$
        !Status = 9
        !ValAbsComp = 0
        !Idebecomp = 0
        !IHABECOMP = 0
        !BASEIMPO = 0
        .Update
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
        GoTo 20
      End If
    End With
    '
    If CCC$ <> "" Then
      CCC$ = Left$(CCC$, 10) + "*"
      With CajaBanc
        .FindFirst "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
22      If .NoMatch = False Then
          .Edit
          !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
          !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
          !DESTICHEQ = ""
          !REFESALCHE = ""
          !IDENSALCHE = ""
          !FEDEPOSENT = CVDAT(0) 'fecha de deposito
          .Update
          .FindNext "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
          GoTo 22
        End If
      End With
    End If
    '
    With Sacrepen
      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
25     If .NoMatch = False Then
        .Delete
        .FindNext "CodiCom_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 25
      End If
    End With
    
    ' ANULA RETENCION DE INGRESOS BRUTOS
    For U& = 1 To ULTREG("SURETIB")
      Z1$ = REGLEIDO$("SURETIB", U&)
      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
        Call REPLA(Z1$, String$(16, 0), 79, 16)
        Call GRAREG("SURETIB", Z1$, U&)
      End If
    Next U&
    
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie = 0
    NumeProv% = NC%
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    Set Aplipag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With Aplipag1
      .FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
35     If .NoMatch = False Then
        CoComCor$ = !CODOCOR_COR
        CoComLar$ = !codocor_lar
        FechEmi = Now
        IBruTot# = !ITOTORIG
        IIvaTot# = !IIVAORIG
        NuorVen% = 0
        FeVenc = Now
        ImVenc# = !IMPOAPLI
        '
        CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        If CajaBanc.NoMatch = False Then
          FechEmi = CajaBanc!FECHEMISI
          FeVenc = CajaBanc!FECHEMISI
        End If
        '
        Call GRACREPEN
        .FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        GoTo 35
      End If
    End With
    '
    Screen.MousePointer = 1
    Call COMUNI("ANULACIóN COMPLETA.\  \No Fue Posible Reconstruir\las Fechas de Vencimiento Originales.")
    RECUEP07.Text1 = TRIM$(Str$(NC%))
    RECUEP07.Show 1
    '
    Call CIERRARCH("CUECORR")
    Call CIERRARCH("PROVED2")
    '
    GoTo 10
    '
End Sub
'
Sub ANUORPA2()       ' anulacion de ordenes de pago
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              '
    
    If DERACCE%("ANUORPA", "Anulación de Ordenes de Pago") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
    'jandy sql
    'FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    FECHALU = FETEXCOR1$(FECHALIM%)
    Dim MIRS As ADODB.Recordset
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    Dim CABATEMP As ADODB.Recordset
    Condi$ = " (TipoMovim = 'OPAG' OR (TipoMovim = 'AJUP' AND OpciMovim = 'CC')) "
    Condi$ = Condi$ + " AND NuProve = " & NC% & " "
    Condi$ = Condi$ + " AND FechEmisi > '" & FECHALU & "' "
    Condi$ = Condi$ + " AND (Left(Codicom_lar, 2) = 'OP' Or Left(Codicom_lar, 2) = 'CC')"
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE " + Condi$
    SQLX$ = SQLX$ + " Order by CODICOM_LAR ASC "

    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$), REGAF)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NUPROVE
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then ' PARA QUE TAMBIEN TRAIGA LAS ANULADAS
'          If UCase$(Left$(!Codicom_lar, 2)) = "OP" Or UCase$(Left$(!Codicom_lar, 2)) = "CC" Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            Y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
'          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Comprobantes de Pago que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
12  Call SELECHO("!SELCUECO")
    If TRIM$(VALACT2$("!SELCUECO")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUECO"), 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6)) ' para anular retenciones i.brutos
    '
    If VERICONCILIA%(DOCULA$) < 1 Then GoTo 10
    '
    ' VERIFICA SI TIENE APLICACIONES PENDIENTES
    Condi$ = "CODOCAP_LAR='" & DOCULA$ & "'"
    REGI& = CantRegistros("APLIPAGO", Condi$)
    If REGI& > 0 Then
      Call COMUNI("Imposible Anular Comprobante.\Debe Primero Desaplicar el Comprobante y Luego Re-intentar Anularlo")
      GoTo 12
    End If
    '
    MODOANU% = COMALTER%("Modo de Anulación del Comprobante '" + DOCULA$ + "'\\Marca de Anulación\Re-Utilizar Numero")
    If MODOANU% < 1 Or MODOANU% > 2 Then GoTo 12
    '
    CCC$ = ""
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
    'Set CajaBanc1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
    'On Error Resume Next
    'CajaBanc1.MoveFirst
    'If Err Then
    '  On Error GoTo 0
    '  Call MENSERR(24, "Imposible Anular la O/Pago Elegida")
    '  GoTo 10
    'End If
    If (.EOF And .BOF) Then 'si no la encuentra
      Call MENSERR(24, "Imposible Anular la O/Pago Elegida")
      GoTo 10
    Else
      'RECUPERO COLUMNAS
      CCC$ = Left$(TRIM$(MIRS!CODICOM_COR), 9) + "-nl"
      CCCX$ = Left$(TRIM$(MIRS!CODICOM_COR), 10) + "*"
      'DATOS PARA LA RECONSTRUCCION
'      CodiEmp% = !CodiEmpr
'      CodiSuc% = !CODISUCU
      CCL$ = !CODICOM_LAR
      Call REPLA(CCL$, "anul", 6, 4)
    End If
    End With
    MIRS.Close
    Set MIRS = Nothing
    
    '
    'SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCCX$ & "' "
    'Set CajaBAnc2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)

    'SQLX$ = "SELECT * FROM SACREPEN "
    'SQLX$ = SQLX$ + "WHERE CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
    'Set SACREPEN2 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
'aca borro
 
 If MODOANU% = 2 Then
'    strsql = "Delete CajaBanc where CodiCom_Lar = '" & DOCULA$ & "'"
'    strsql = strsql & " AND NuProve = " & NC%
'    Call EJECUTACOMANDO(strsql)
    Condi$ = "CodiCom_Lar = '" & DOCULA$ & "'"
    Condi$ = Condi$ + " and NuProve = " & NC%
    CANTIX = CantRegistros("CAJABANC", Condi$, "NOMESS")
    
    Call BORRAREGISTROS("CajaBanc", Condi$, REG&)
    
    GoTo 21
 End If
    
'    With CajaBAnco
'19    .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'20    If .NoMatch = False Then
'        ' DATOS PARA LA RECONSTRUCCION
'          CodiEmp% = !CodiEmpr
'          CodiSuc% = !CodiSucu
'        '
'        If MODOANU% = 2 Then
'           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
'          .Delete
'          GoTo 19
'        End If
'        '
'        .Edit
'           CCL$ = !codicom_lar
'           Call REPLA(CCL$, "anul", 6, 4)
'           !codicom_lar = CCL$
'           CCC$ = Left$(TRIM$(!CodiCom_Cor), 9) + "-nl"
'           !CodiCom_Cor = CCC$
'           !Status = 9
'           !ValAbsComp = 0
'           !IDEBECOMP = 0
'           !IHABECOMP = 0
'           !BASEIMPO = 0
'          .Update
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & ""
'        GoTo 20
'       End If
'    End With
    '
     
'aca actualizo
      strsql = "Update CajaBanc set codicom_lar = '" & CCL$ & "'"
      strsql = strsql & ", CodiCom_Cor = '" & CCC$ & "'"
      strsql = strsql & ", Status = 9 "
      strsql = strsql & ", ValAbsComp = 0 "
      strsql = strsql & ", IDEBECOMP = 0 "
      strsql = strsql & ", IHABECOMP = 0 "
      strsql = strsql & ", BASEIMPO = 0 "
      strsql = strsql & " WHERE CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC%
      Call EJECUTACOMANDO(strsql)
       
'aca ACTUALIZO SEGUN CCC$
21  If CCC$ <> "" Then
      'CCC$ = Left$(CCC$, 10) + "*"
      CCC$ = Left$(CCC$, 10) + "%"
      strsql = "Update CajaBanc set Status=1 "
      strsql = strsql & ", Document =''"
      strsql = strsql & ", DESTICHEQ =''"
      strsql = strsql & ", REFESALCHE = ''"
      strsql = strsql & ", IDENSALCHE = ''"
      strsql = strsql & ", FEDEPOSENT = '" & FETEXCOR1$(33) & "'"
      strsql = strsql & ", FECHENTRE = '" & FETEXCOR1$(33) & "'"
      strsql = strsql & " WHERE TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
      Call EJECUTACOMANDO(strsql)
    End If

'      With CajaBAnc2
'        .FindFirst "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
'22      If .NoMatch = False Then
'          .Edit
'          !Status = 1     'LO PONGO DEVUELTA COMO PENDIENTE EN CARTERA DE VALORES
'          !Document = ""  'QUITO EL NUMERO DE DEPOSITO ASIGNADO
'          !DESTICHEQ = ""
'          !REFESALCHE = ""
'          !IDENSALCHE = ""
'          !FEDEPOSENT = CVDAT(0) 'fecha de deposito
'          !FECHENTRE = CVDAT(0)  'fecha de ENTREGA
'          .Update
'          .FindNext "TipoMovim = 'CHRECI' AND DocumEnt LIKE '" & CCC$ & "' "
'          GoTo 22
'        End If
'      End With
'    End If
    '

    
'    With SACREPEN2
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'25     If .NoMatch = False Then
'        .Delete
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
'        GoTo 25
'      End If
'    End With
    '
    GoTo 80
' aca es delete en SACREPEN
'     strsql = "Delete SACREPEN where CodiCom_Lar = '" & DOCULA$ & "'"
'     strsql = strsql & " AND NuProve = " & NC%
     Condi$ = "CodiCom_Lar = '" & DOCULA$ & "'"
     Condi$ = Condi$ + " AND NuProve = " & NC%
     Call BORRAREGISTROS("SACREPEN", Condi$, REG&, "NOMESS")
     
' aca es delete en ApliPAgo
'     strsql = strsql & " Delete APLIPAGO where CodocAp_lar = '" & DOCULA$ & "'"
'     strsql = strsql & " AND NuProve = " & NC%
'     Call EJECUTACOMANDO(strsql)
      Condi$ = "CodocAp_lar = '" & DOCULA$ & "'"
      Condi$ = Condi$ + " AND NuProve = " & NC%
      Call BORRAREGISTROS("APLIPAGO", Condi$, REG&, "NOMESS")
     
'ESTO SE COMENTA Y SE PASA ABAJO PARA QUE QUITE EL IMPORTE DEL DAT SURETIB
'
'    For U& = 1 To ULTREG("SURETIB")
'      Z1$ = REGLEIDO$("SURETIB", U&)
'      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
'        Call REPLA(Z1$, String$(16, 0), 79, 16)
'        Call GRAREG("SURETIB", Z1$, U&)
'      End If
'    Next U&
    '
    ' RECONSTRUIR COMPOSICION DE DEUDA
    NumeClie = 0
    NumeProv% = NC%
    Dim Aplipag1 As ADODB.Recordset
    SQLX$ = "SELECT * FROM ApliPago "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    Set Aplipag1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'Set Aplipag1 = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With Aplipag1
      If Not (.EOF And .BOF) Then
        .MoveFirst
'.FindFirst "CoDocAp_Lar = '" & DOCULA$ & "' AND NuProve = " & NC% & " "
'35     If .NoMatch = False Then
        Do While Not (.EOF)
          CoComCor$ = !CODOCOR_COR
          CoComLar$ = !codocor_lar
          FechEmi = Now
          IBruTot# = !ITOTORIG
          IIvaTot# = !IIVAORIG
          NuorVen% = 0
          FeVenc = Now
          ImVenc# = !IMPOAPLI
        '
        'CajaBanc.FindFirst "CodiCom_Lar = '" & CoComLar$ & "' AND NuProve = " & NC% & " "
        'If CajaBanc.NoMatch = False Then
          strsql = "Select * from CajaBanc "
          strsql = strsql & "Where CodiCom_Lar = '" & CoComLar$ & "'"
          strsql = strsql & " AND NuProve = " & NC%
          Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
          With MIRS
            If Not (.EOF And .BOF) Then
                FechEmi = !FECHEMISI
                FeVenc = !FECHEMISI
            End If
          End With
          MIRS.Close
          Set MIRS = Nothing
        'End If
        '
        Call GRACREPEN
        '.FindNext "CoDocAp_Lar = '" & DOCULA$ & "'AND NuProve = " & NC% & ""
        'GoTo 35
        .MoveNext
        Loop
      End If
    End With
    Aplipag1.Close
    Set Aplipag1 = Nothing
    '
    '
    Screen.MousePointer = 1
    
    Call COMUNI("ANULACIóN COMPLETA.\  \No Fue Posible Reconstruir\las Fechas de Vencimiento Originales.")
    RECUEP07.Text1 = TRIM$(Str$(NC%))
    RECUEP07.Show 1
    '
80
    'CODIGO TRASLADADO AQUI PARA QUE LIMPIE DEL DAT EL IMPORTE
    For U& = 1 To ULTREG("SURETIB")
      Z1$ = REGLEIDO$("SURETIB", U&)
      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
        Call REPLA(Z1$, String$(16, 0), 79, 16)
        Call GRAREG("SURETIB", Z1$, U&)
      End If
    Next U&
    
    ' ANULA RETENCION DE GANANCIAS
    For U& = 1 To ULTREG("SUREGAN")
      Z1$ = REGLEIDO$("SUREGAN", U&)
      If CVS(Mid$(Z1$, 7, 4)) = NUOPANU& Then
        Call REPLA(Z1$, String$(16, 0), 79, 16)
        Call GRAREG("SUREGAN", Z1$, U&)
      End If
    Next U&

    Call CIERRARCH("CUECORR")
    Call CIERRARCH("PROVED2")
    '
    Call RAPLIPRO(NC%)
    '
    If MODOANU% = 2 Then 'solo en el caso que elije reutilizar numero
      Screen.MousePointer = 11
      Call ORPAGO07.ACTULISTAOP ' ACTUALIZA LA LISTA DE SELECCION
      Screen.MousePointer = 1
    End If

    GoTo 10
    '
End Sub
'
Private Sub Form_Load()
   If EXISTE%("ANAPAGOS.EXE") < 1 Then
     Command9.Visible = False
   End If
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture7)
End Sub

Sub AJUSALP()
    '
    Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%)
    Screen.MousePointer = 1
    CLSS$ = OBJPLA$("SELPROAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    '
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = rasocli$((-1) * NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(Control$("", "CIEMECOM"))
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
    'JANDY SQL
    Dim CABATEMP As ADODB.Recordset
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "or TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUP') "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
        IFAC# = !IHABECOMP - !Idebecomp
        FF% = CVINT(!fechconta)
        SALFINAL# = SALFINAL# + IFAC#
        If FF% <= FAJUI% Then
          SARRASFE# = SARRASFE# + IFAC#
        End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
'3.- Determina saldo por composicion de deuda
    'SQLX$ = "SELECT * FROM SACREPEN "
    'SQLX$ = SQLX$ + "where NuProve = " & NC% & " "
    'Set CODEUTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    
    strsql = "Select Sum(SaldAcre - SALDDebe) as SALDEUTO from SACREPEN"
    strsql = strsql & " where NuProve = " & NC%
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
    If Not (MIRS.EOF And MIRS.BOF) Then
      SALDEUTO# = ROUND#(MIRS!SALDEUTO)
    End If
    MIRS.Close
    Set MIRS = Nothing
    
    
  '
'    With CODEUTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err > 0 Then
'        On Error GoTo 0
'        GoTo 100
'      End If
'      On Error GoTo 0
'      Do While Not .EOF
'        SALDEUTO# = SALDEUTO# + !SALDAcre - !SaldDebe
'      .MoveNext
'      Loop
'    End With
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0)

102 CLSS$ = OBJPLA$("CARAJUPRO", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 13, 8))
    ICONI% = CVI(Mid$(CLSS$, 21, 2))
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
      Call MENSERR(24, "Falta Imputación Contable Ajuste")
      VDEF$ = CLSS$
      GoTo 102
    End If
    If ICONI% = CUECOCLI%((-1) * NC%) Then
      Call MENSERR(24, "Cuenta Contable no Válida")
      VDEF$ = CLSS$
      GoTo 102
    End If
    '
    DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
    DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
    '
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    '
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
'    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
'    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    '
    ORMA% = (-1)
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
    End If
    RECITEMP.Close
    Set RECITEMP = Nothing
    
'    Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With RECITEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
'      End If
'    End With
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
20  NC1% = (-1) * NC%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FAJUI%)
    '
    Referen$ = rasocli$(NC1%)
    FeContab = CVDAT(FAJUI%)
    FeAcredi = CVDAT(FAJUI%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = rasocli$(NC1%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1%))
    TiMovi = "AJUP"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALPR"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + rasocli$(NC1%)
    ImpoDeb# = DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FAJUI%)
    ImVenc# = DIFEDEU#
    Call GRACREPEN
    '
    'CajaBanc.Close
    Call COMUNI("Ajuste Completado")
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'
Sub VAPLIOP(Optional CLIACTIVO%)
  '\\\OT-05-0441-WAR-30/06/05///
10  Screen.MousePointer = 1
    '\\\OT-07-0190-OD-24/04/08///
    If CLIACTIVO% > 0 Then NC% = CLIACTIVO%: GoTo 20
    '\\\OT-07-0190-OD-FIN///

    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
20  If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
    Dim APCOTMP As ADODB.Recordset
    SQLX$ = " SELECT  CoDocAp_Cor, sum(ImpoApli) as TOTAPLI FROM APLIPAGO "
    SQLX$ = SQLX$ + " WHERE NuPROVE = " & NC% & ""
    SQLX$ = SQLX$ + " GROUP BY CoDocAp_Cor"
    SQLX$ = SQLX$ + " ORDER BY CoDocAp_Cor ASC "
    Set APCOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With APCOTMP
      If (.EOF And .BOF) Then
        Call MENSERR(24, "Sin Movimientos")
        If CLIACTIVO% > 0 Then GoTo 999 'SI CLIENTE SELECCIONADO PREVIO ,SALE
        GoTo 10
      End If '
      
      .MoveFirst
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      Do While Not .EOF
           JJ& = JJ& + 1
         Call REPLA(Y$, MKS(JJ&), 1, 4)
           NROX$ = !CODOCAP_COR
         Call REPLA(Y$, NROX$, 5, 18)
           IMPOS$ = NUMSTRI$(!TOTAPLI, 13, 2, 1)
         Call REPLA(Y$, IMPOS$, 23, 13)
         Call GRAREG("!SELCUERE", Y$, JJ&)
     .MoveNext
      Loop
      
      Call SETULTREG("!SELCUERE", JJ&)
      '
    End With
    APCOTMP.Close
    Set APCOTMP = Nothing
    
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(rasocli$(NC% * -1)))
    
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + rasocli$(NC%)
      '
      JJ& = 0
      
      Dim Aplipago As ADODB.Recordset
      strsql = "Select * from Aplipago where NuPROVE = " & NC%
      strsql = strsql & "And  CoDocAp_Cor = '" & TRIM(DAPLI$) & "'"
      Set Aplipago = FLEXCONN.Execute(FILTSQL$(strsql))
      With Aplipago
        '
        If Not (.EOF And .BOF) Then
          .MoveFirst
          Do While Not (.EOF)
           Y$ = REGBLAN$("!SELDOCAP")
             NROY$ = !codocor_lar
             REFY$ = Left(!refeapli, 18)
             If UCase(Left(REFY$, 5)) = "ORDEN" Then
                 REFY$ = Left(!refeapli, 14)
             End If
             IMPO# = !IMPOAPLI
             IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
           Call REPLA(Y$, NROY$, 1, 18)
           Call REPLA(Y$, REFY$, 19, 24)
           Call REPLA(Y$, IMPOS$, 44, 13)
           JJ& = JJ& + 1
           Call GRAREG("!SELDOCAP", Y$, JJ&)
           .MoveNext
           Loop
        End If
      End With
      Aplipago.Close
      Set Aplipago = Nothing
      
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    If CLIACTIVO% > 0 Then GoTo 999 'SI CLIENTE SELECCIONADO PREVIO ,SALE
    '
    GoTo 10
    '
999 Screen.MousePointer = 1
    End Sub

Sub VAPLIFAC()
    '
10  Screen.MousePointer = 1
    Call SELECHO("PROVED1")
    NC% = XVALO(VALACT1$("PROVED1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
    'jandy sql Se uso una unica select
    Call APERBASDAT("CABAN")
    Dim APCOTMP As ADODB.Recordset
    SQLX$ = " SELECT  DISTINCT CoDocOr_Cor, ImpoApli  FROM APLIPAGO "
    SQLX$ = SQLX$ + " WHERE NuPROVE = " & NC% & ""
    SQLX$ = SQLX$ + " ORDER BY CoDocOr_Cor ASC "
    Set APCOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With APCOTMP
      If (.EOF And .BOF) Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If '
      '
      .MoveFirst
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0: NROX$ = "": IMPO# = 0
      CODVIE$ = ""
      Do While Not .EOF
          If !CODOCOR_COR = NROX$ Then
             IMPO# = IMPO# + !IMPOAPLI
             IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
             Call REPLA(Y$, IMPOS$, 23, 13)
             Call GRAREG("!SELCUERE", Y$, JJ&)
             GoTo 19
          End If
          NROX$ = !CODOCOR_COR
          Call REPLA(Y$, NROX$, 5, 18)
          IMPO# = !IMPOAPLI
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          J1% = Int(JJ&)
          Call REPLA(Y$, MKS(J1%), 1, 4)
          Call GRAREG("!SELCUERE", Y$, JJ&)
19    .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
'    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(rasocli$(-1 * NC%)))
    
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + rasocli$(-1 * NC%)
      '
      JJ& = 0
      'jandy sql esto deberia ser select sum()....group by
      Dim Aplipago As ADODB.Recordset
      strsql = "Select DISTINCT CoDocAp_Lar, RefeApli, ImpoApli from Aplipago where NuPROVE = " & NC%
      strsql = strsql & "And  CoDocOr_Cor = '" & TRIM(DAPLI$) & "'"
      Set Aplipago = FLEXCONN.Execute(FILTSQL$(strsql))
      With Aplipago
         If Not (.EOF And .BOF) Then
            .MoveFirst
            Y$ = REGBLAN$("!SELDOCAP")
            Do While Not (.EOF)
               NROY$ = !codocap_lar
               REFY$ = Left(!refeapli, 18)
               If TRIM$(REFY$) = "" Then REFY$ = "." + REFY$
               IMPO# = !IMPOAPLI
               IMPOS = FORMATNUM$(IMPO#, "F13.2")
               '
               'CALCULO EL TOTAL
               IMPO# = 0
               Dim MIRS1 As ADODB.Recordset
               strsql1 = "Select DISTINCT ImpoApli from Aplipago"
               strsql1 = strsql1 & " where NuPROVE = " & NC%
               strsql1 = strsql1 & " And  CoDocAP_lar = '" & NROY$ & "' "
               Set MIRS1 = FLEXCONN.Execute(FILTSQL$(strsql1))
               With MIRS1
               If Not (.EOF And .BOF) Then
                  .MoveFirst
                  Do While Not (.EOF)
                   IMPO# = IMPO# + !IMPOAPLI
                   .MoveNext
                  Loop
               End If
               End With
               MIRS1.Close
               Set MIRS1 = Nothing
               '
               Y$ = Space$(60)
               Call REPLA(Y$, NROY$, 1, 18)
               Call REPLA(Y$, REFY$, 19, 24)
               Call REPLA(Y$, IMPOS$, 44, 13)
               JJ& = JJ& + 1
               Call GRAREG("!SELDOCAP", Y$, JJ&)
            .MoveNext
            Loop
         End If
         '
      End With
      Aplipago.Close
      Set Aplipago = Nothing
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
  SQLX1$ = SQLX1$ + " WHERE CodiCom_Lar = '" & DOC1$ & "'"
  'Set NROTMP = CueCorri.OpenRecordset(SQLX1$, 4)
  Dim NROTMP As ADODB.Recordset
  Set NROTMP = New ADODB.Recordset
24 On Error Resume Next
   NROTMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
   End If
   On Error GoTo 0

  '
  SQLX$ = " SELECT * FROM APLIPAGO "
  'Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  Dim DOCAPTMP As ADODB.Recordset
  Set DOCAPTMP = New ADODB.Recordset
26 On Error Resume Next
   DOCAPTMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
  '
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
    YY$ = REGBLAN("APLIPAG")
    FEC1% = 0
    CODVIE$ = ""
    .MoveFirst
    Do While Not .EOF
     'If CODVIE$ <> !CoDocAP_LAR Then '***
      Call REPLA(YY$, MKI(!NUPROVE), 3, 2)
      Call REPLA(YY$, !codocap_lar, 5, 18)
      Call REPLA(YY$, MKD(!IMPOAPLI), 23, 8)
      Call REPLA(YY$, Left(!refeapli, 24), 31, 24)
      Call REPLA(YY$, !codocor_lar, 55, 18)
       DOC1$ = !codocap_lar
       With NROTMP
         On Error Resume Next
         '.FindFirst "CodiCom_Lar = '" & DOC1$ & "' "
         'If .NoMatch = True Then
         If (.EOF And .BOF) Then
          On Error GoTo 0
          GoTo 25
         End If
         On Error GoTo 0
          FEC1% = CVINT(!FECHEMISI)
       End With
       
25    Call REPLA(YY$, MKI(FEC1%), 1, 2)
      JJ& = JJ& + 1
      Call TRACE(24, JJ&, Max&)
      Call GRAREG("APLIPAG", YY$, JJ&)
      'CODVIE$ = !CoDocAP_LAR
      'End If
      .MoveNext
    Loop
    '
  End With
  NROTMP.Close
  DOCAPTMP.Close
  Set DOCAPTMP = Nothing
  Set NROTMP = Nothing
  Call SETULTREG("APLIPAG", JJ&)
  Call CIERRARCH("*.*")
  '
End Sub
'\\\OT-05-0441-WAR-FIN///
Sub APLIFEC()
   '
   Call DESHASFECHA(DES%, Has%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   DESA = FETEXCOR1$(DES%)
   HASA = FETEXCOR1$(Has%)
   '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT NUPROVE, FechEmisi, CodiCom_Lar, ValAbsComp FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & DESA & "'  "
    SQLX$ = SQLX$ + " AND FechEmisi <= '" & HASA & "'  "
    SQLX$ = SQLX$ + " AND ( TipoMovim = 'OPAG' OR (TipoMovim = 'FCOM'AND OpciMovim = 'ND' ) ) "
    'SQLX$ = SQLX$ + " AND TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + " Order by FechEmisi ASC "
    'Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
    Dim NROTMP As ADODB.Recordset
    Set NROTMP = New ADODB.Recordset
25  On Error Resume Next
    NROTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0

    '
    SQLX$ = " SELECT * FROM APLIPAGO "
    SQLX$ = SQLX$ + " CODOCAP_LAR = '" & DOCAPLI$ & "'"
    'Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
    Dim DOCAPTMP As ADODB.Recordset
    Set DOCAPTMP = New ADODB.Recordset
28  On Error Resume Next
    DOCAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 28
    End If
    On Error GoTo 0

    With NROTMP
      On Error Resume Next
      .MoveFirst
      If Err Then GoTo 99
      On Error GoTo 0
      Do While .EOF = False
        YY$ = REGBLAN("APLIPAG")
        FEDOC% = CVINT(!FECHEMISI)
        NCLIE% = !NUPROVE
        DOCAPLI$ = !CODICOM_LAR
        '
        With DOCAPTMP
         On Error Resume Next
         'SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
         '.FindFirst SQLX2$
  '36     'If DOCAPTMP.NoMatch = False Then
         If Not (.EOF And .BOF) Then
           On Error GoTo 0
           Do While Not (.EOF)
            Call REPLA(YY$, MKI(FEDOC%), 1, 2)
            Call REPLA(YY$, MKI(NCLIE%), 3, 2)
            Call REPLA(YY$, DOCAPLI$, 5, 18)
              IAPLI# = !IMPOAPLI
            Call REPLA(YY$, MKD(!IMPOAPLI), 23, 8)
              REFEA$ = Left(!refeapli, 18)
            Call REPLA(YY$, REFEA$, 31, 24)
            Call REPLA(YY$, !codocor_lar, 55, 18)
            JJ& = JJ& + 1
            'Call TRACE(24, JJ&, Max&)
            Call GRAREG("APLIPAG", YY$, JJ&)
            '.FindNext SQLX2$
            'GoTo 36
            .MoveNext
           Loop
         End If
         On Error GoTo 0
        End With
26    .MoveNext
      Loop
      '
  End With
  NROTMP.Close
  DOCAPTMP.Close
  Set NROTMP = Nothing
  Set DOCAPTMP = Nothing
90 Screen.MousePointer = 1
    Call SETULTREG("APLIPAG", JJ&)
    Call HEADERS("APLIPAG", "")
    Call HEADERS("APLIPAG", "Corresponde a: " + TRIM$(PERIO$))
    Call FINAL("*LIAPLIPA")
    '
99  On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

'\\\OT-06-0405-MH-22/09/06///
Sub LIAPLIFE()
   '
   Dim FEDOCAP%(), NPROVED%(), NUDOCAP$(), IMPORT#()
   '
   Call DESHASFECHA(DES%, Has%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   'DESA = Int(CDbl(CVDAT(DES%)))
   'HASA = Int(CDbl(CVDAT(HAS%)))
   'jandy sql
   
   DESA = FETEXCOR1$(DES%)
   HASA = FETEXCOR1$(Has%)
   
   '
   Call APERBASDAT("CABAN")
   Dim NROTMP As ADODB.Recordset
   SQLX$ = " SELECT NUPROVE, IDebeComp,IHabeComp, FechEmisi, CodiCom_Lar FROM CAJABANC "
   'SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & " "
   'SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & " "
   SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & DESA & "' "
   SQLX$ = SQLX$ + " AND FechEmisi <= '" & HASA & "' "
   SQLX$ = SQLX$ + " AND ( TipoMovim = 'OPAG' OR (TipoMovim = 'FCOM'AND OpciMovim = 'NC' ) ) "
   SQLX$ = SQLX$ + " Order by FechEmisi ASC, CodiCom_Lar ASC "
   'Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
   Set NROTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   CADOC% = 0: DOCANT$ = ""
   With NROTMP
'     On Error Resume Next
'     .MoveFirst
'     If Err Then GoTo 90
'     On Error GoTo 0
'     '
   If Not (.EOF And .BOF) Then
     .MoveFirst
     Do While .EOF = False
       FEDOC% = CVINT(!FECHEMISI)
       NPROVE% = !NUPROVE
       DOCAPLI$ = !CODICOM_LAR
       IMPOTO# = !Idebecomp - !IHABECOMP
       If Abs(IMPOTO#) < 0.005 Then GoTo 26
       '
       If DOCAPLI$ = DOCANT$ Then
         IMPORT#(CADOC%) = IMPORT#(CADOC%) + IMPOTO#
         GoTo 26
       End If
       '
       CADOC% = CADOC% + 1
       ReDim Preserve FEDOCAP%(CADOC%), NPROVED%(CADOC%), NUDOCAP$(CADOC%), IMPORT#(CADOC%)
       FEDOCAP%(CADOC%) = FEDOC%
       NPROVED%(CADOC%) = NPROVE%
       NUDOCAP$(CADOC%) = DOCAPLI$
       IMPORT#(CADOC%) = IMPOTO#
       DOCANT$ = DOCAPLI$
       '
26   .MoveNext
     Loop
     '
   End If
   End With
   NROTMP.Close
   Set NROTMP = Nothing
   '
   'JANDY SQL
   'SQLX$ = " SELECT * FROM APLIPAGO "
   'Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
For IDOC% = 1 To CADOC%
      IMPOTO# = IMPORT#(IDOC%)
      RESTO# = IMPOTO#
      '
      PRIREN% = 1
      DOCAPLI$ = NUDOCAP$(IDOC%)
      NPROVE% = NPROVED%(IDOC%)
      FEDOC% = FEDOCAP%(IDOC%)
      
   '   SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
   '   DOCAPTMP.FindFirst SQLX2$
      '
'36    If DOCAPTMP.NoMatch = False Then
         
   Dim DOCAPTMP As ADODB.Recordset
   Set DOCAPTMP = New ADODB.Recordset
   DOCAPTMP.CursorLocation = adUseClient
   SQLX$ = "Select * from APLIPAGO"
   SQLX$ = SQLX$ & " where CODOCAP_LAR = '" & DOCAPLI$ & "' "
25  On Error Resume Next
    DOCAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    With DOCAPTMP
    If Not (.EOF And .BOF) Then
     .MoveFirst
     Do While Not .EOF()
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         'AGREGUE(MH)
         'Reemplazaba por 'NCLIE%
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
           Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
           PRIREN% = 0
         End If
         Call REPLA(YY$, DOCAPTMP!codocor_lar, 31, 18)
           REFEA$ = Left(DOCAPTMP!refeapli, 24)
         Call REPLA(YY$, REFEA$, 49, 24)
           IAPLI# = DOCAPTMP!IMPOAPLI
           If IAPLI# > RESTO# Then IAPLI# = RESTO#
         Call REPLA(YY$, MKD(IAPLI#), 73, 8)
         RESTO# = RESTO# - IAPLI#
         If IAPLI# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("APLIPA1", YY$, JJ&)
           CAREGRA% = CAREGRA% + 1
         End If
'         DOCAPTMP.FindNext SQLX2$
'         GoTo 36
          .MoveNext
          Loop
      End If
    End With
      
      '
      ' ESTO ES PARA PREVER QUE ESA o/pAGO AUN NO SE APLICó NADA
      If Abs(RESTO#) >= 0.005 Then
         ' buscar si aun está pendiente de aplicacion
         REFEB$ = "*** Pend.Aplic.***"
         
         Dim Sacrepen As ADODB.Recordset
         strsql = "Select * from SacrePen where NuProve= " & NPROVE%
         strsql = strsql & " And CodiCom_Lar = '" & DOCAPLI$ & "'"
         strsql = strsql & " And (SaldDebe + SaldAcre) > 0.005 "
         Set Sacrepen = FLEXCONN.Execute(FILTSQL$(strsql))
         With Sacrepen
         'Sacrepen.FindFirst "NuProve = " & NPROVE% & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         'If Sacrepen.NoMatch = True Then
         If (.EOF And .BOF) Then 'si no lo encuentra
            REFEB$ = "Aplic.Automatica"
            NC1% = (-1) * NPROVE%
            If TICUEPRO%(NC1%) = 1 Then ' CUENTA DE GASTOS
              REFEB$ = "Cuenta de Gastos"
            End If
         End If
         End With
         Sacrepen.Close
         Set Sacrepen = Nothing
         '
         YY$ = REGBLAN("APLIPA1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKI(NPROVE%), 21, 2)
         If PRIREN% = 1 Then
            Call REPLA(YY$, MKD(IMPOTO#), 23, 8)
         End If
         Call REPLA(YY$, REFEB$, 31, 18)
           REFEA$ = "Pendiente Aplic."
         Call REPLA(YY$, REFEA$, 49, 24)
         Call REPLA(YY$, MKD(RESTO#), 73, 8)
         JJ& = JJ& + 1
         Call GRAREG("APLIPA1", YY$, JJ&)
    End If
            '
Next IDOC%
   '
90 Call SETULTREG("APLIPA1", JJ&)
   Call HEADERS("APLIPA1", "")
   Call HEADERS("APLIPA1", "Corresponde a: " + TRIM$(PERIO$))
   Call FINAL("*LIAPLIP1")
   '
99 Screen.MousePointer = 1
   '
End Sub
Sub TRANSORPA()

   ' Transferencia de Cuenta Corriente
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
    '\\\OT-08-0814-OD-25/09/08///

     FECHALIM% = XVALO(Control$("", "CIEMECOM"))
     '
10   Screen.MousePointer = 1
     Call SELECHO("PROVED1")
     NC% = XVALO(VALACT1$("PROVED1"))
     If NC% < 1 Then Exit Sub
     '
     Screen.MousePointer = 11
     JJ& = 0
100  Call SETULTREG("!SELCUECO", 0)
     FechaLimite = FETEXCOR1$(FECHALIM%)
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ & " WHERE TipoMovim = 'OPAG'"
     SQLX$ = SQLX$ & " AND NuProve  = " & NC% & ""
     SQLX$ = SQLX$ & " AND FechEmisi  > '" & FechaLimite & "'"
     'SQLX$ = SQLX$ & " ORDER BY  cast(substring(CodiCom_Cor,4,6) as int) "
     SQLX$ = SQLX$ & " ORDER BY  NUMECOMP ASC "
     '
     Dim CAJABTMP As ADODB.Recordset
     Set CAJABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CAJABTMP
       Do While Not .EOF
        JJ& = JJ& + 1
        Call REPLA(x$, !CODICOM_LAR, 7, 18)
        IMPO# = !Idebecomp - IHABECOMP
        If IMPO# >= 0.005 Then
           IMPOS$ = FORMATNUM$(IMPO#, "F12.2")
           Y$ = MKS$(JJ&) + Mid$(x$, 7, 18) + IMPOS$
           Call REGAPP("!SELCUECO", Y$)
        End If
        .MoveNext
       Loop
     End With
     CAJABTMP.Close
     Set CAJABTMP = Nothing
     Screen.MousePointer = 1


 '     MAXDEU& = ULTREG&("CUECORR")
'     RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
'     '
'     For U& = 1 To Len(RFD$) Step 4
'        REG& = CVS(Mid$(RFD$, U&, 4))
'        X$ = REGLEIDO$("CUECORR", REG&)
'        If CVI(Mid$(X$, 5, 2)) <> NC% Then
'            X$ = REGLEIDO$("CUECORR", 1)
'            Call GRAREG("CUECORR", X$, 1)
'            Call CIERRARCH("CUECORR")
'            GoTo 100
'        End If
'        '
'        If CVI(Left$(X$, 2)) > FECHALIM% Then
'           TIDOC$ = UCase$(Mid$(X$, 7, 2))
'           If TIDOC$ = "OP" Then
'              IMPO# = Abs(CVD(Mid$(X$, 115, 8))) + Abs(CVD(Mid$(X$, 123, 8)))
'              If IMPO# >= 0.005 Then
'                 IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
'                 Y$ = MKS$(REG&) + Mid$(X$, 7, 18) + IMPOS$
'                 Call REGAPP("!SELCUECO", Y$)
'              End If
'           End If
'        End If
'     Next U&
     '
     Call CIERRARCH("!SELCUECO")
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Ordenes de Pago que Puedan Transferirse.")
        GoTo 10
     End If
     '
     Call FRESHECHO("!SELCUECO")
     '
     Call SELECHO("!SELCUECO")
     If VALACT1$("!SELCUECO") = "" Then GoTo 10
     DOCUNU$ = TRIM(Mid$(VALACT2$("!SELCUECO"), 1, 18))
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ & " WHERE TipoMovim = 'OPAG'"
     SQLX$ = SQLX$ & " AND NuProve  = " & NC% & ""
     SQLX$ = SQLX$ & " AND CodiCom_Lar  = '" & DOCUNU$ & "'"
     
     Dim CAJABTMP1 As ADODB.Recordset
     Set CAJABTMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CAJABTMP1
        IMPO# = XVALO(!Idebecomp - IHABECOMP)
        FEDOC% = CVINT(!FECHEMISI)
        NCX% = XVALO(!NUPROVE)
     End With
     '
     Screen.MousePointer = 1

     VDEF$ = DOCUNU$ + MKI$(FEDOC%) + MKD$(IMPO#) + MKI$(NCX%) + MKI$(0)
     OBX$ = OBJPLA$("TRANSORPA", VDEF$)
     '
     

     If OBX$ <> "" Then
       NCY% = CVI(Mid$(OBX$, 31, 2))
       If NCY% > 0 Then
         If NCY% <> NCX% Then
           If COMALTER%("Confirma Transferir: " + DOCUNU$ + "\  de la Cuenta: " & NCX% & " a la Cuenta: " & NCY% & " \\Abandona\Confirma Transferencia") <> 2 Then GoTo 10
           '
           CueMaOriginal$ = TRIM$(CUEMADRE$((-1) * NCX%))
           CueMaDestino$ = TRIM$(CUEMADRE$((-1) * NCY%))
           If CueMaOriginal$ = CueMaDestino$ Then
           '
'            Call REPLA(XX0$, MKI$(NCY%), 5, 2)
'            Call GRAREG("CUECORR", XX0$, REGCUEC&)
'            Call CIERRARCH("CUECORR")
'            Call CIERRARCH("*.*")
'            Call BAJALDISCO
            Screen.MousePointer = 11
            RazSoc$ = Left$(TRIM$(rasocli$((-1) * NCY%)), 48)
            SQLX$ = "Update CajaBanc set NuProve =" & NCY% & ""
            SQLX$ = SQLX$ + " , RefeCom = '" & RazSoc$ & "'"
            SQLX$ = SQLX$ & " WHERE TipoMovim = 'OPAG'"
            SQLX$ = SQLX$ & " AND CodiCom_Lar  = '" & DOCUNU$ & "'"
            SQLX$ = SQLX$ & " AND NuProve  = " & NCX% & ""
            '
             FLEXCONN.Execute (FILTSQL$(SQLX$))
            '
            Screen.MousePointer = 1
            '
            Call COMUNI("Transferencia Completada.\ \FLEXOFT Debe Ahora Reconstruir\la Composición de Deuda\de las dos Cuentas Implicadas.")
            RECUEP07.Text1.TEXT = TRIM$(Str$(NCX%))
            RECUEP07.Show 1
            RECUEP07.Text1.TEXT = TRIM$(Str$(NCY%))
            RECUEP07.Show 1
           Else
            Call MENSERR(24, "Imposible Realizar Transferencia , Difiere Cuenta Madre\del Proveedor: " & NCX% & "  y el Proveedor: " & NCY%)
           End If
         End If
       End If
     End If
     '
    '\\\OT-08-0814-OD-FIN///

End Sub


