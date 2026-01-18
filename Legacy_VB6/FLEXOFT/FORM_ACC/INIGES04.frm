VERSION 5.00
Begin VB.Form INIGES04 
   BackColor       =   &H00C2C2C2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Gestión"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11865
   Icon            =   "INIGES04.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11865
   Begin VB.PictureBox Picture5 
      Height          =   8055
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   10755
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Image Image2 
         Height          =   7990
         Left            =   0
         Picture         =   "INIGES04.frx":1D2A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10815
      End
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
      Height          =   600
      Left            =   2520
      Picture         =   "INIGES04.frx":B9C8
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Expedición - Remito de Ventas - Clientes "
      Top             =   6300
      Visible         =   0   'False
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
      Left            =   9135
      Picture         =   "INIGES04.frx":BCD2
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Sistema de Contabilidad General"
      Top             =   1575
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   315
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   30
      Top             =   2940
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   8610
      Pattern         =   "*.EXK"
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   1590
   End
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
      Left            =   8760
      Picture         =   "INIGES04.frx":BFDC
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Sub-Sistema de Compras Industriales"
      Top             =   6240
      Visible         =   0   'False
      Width           =   600
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
      Left            =   3150
      Picture         =   "INIGES04.frx":C2E6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Sub-Sistema de Pedidos de Clientes"
      Top             =   5460
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   450
   End
   Begin VB.CommandButton Command15 
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
      Left            =   7050
      Picture         =   "INIGES04.frx":C728
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Acreedores - Cuentas por Pagar"
      Top             =   5205
      Width           =   600
   End
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
      Height          =   600
      Left            =   1200
      Picture         =   "INIGES04.frx":CB6A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Gestión de Ventas y Deudores"
      Top             =   5130
      Width           =   600
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
      Height          =   600
      Left            =   7380
      Picture         =   "INIGES04.frx":CE74
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Tesorería y Gestión de Valores"
      Top             =   510
      Width           =   600
   End
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
      Left            =   720
      Picture         =   "INIGES04.frx":D17E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   450
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   4305
      Picture         =   "INIGES04.frx":D488
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   6
      Top             =   6525
      Width           =   1905
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   160
         TabIndex        =   14
         Top             =   450
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   13
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   12
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   11
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8115
      Left            =   10800
      ScaleHeight     =   8055
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox Picture16 
         Height          =   8055
         Left            =   0
         ScaleHeight     =   7995
         ScaleWidth      =   1140
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         Begin VB.Image Image3 
            Height          =   8115
            Left            =   0
            Picture         =   "INIGES04.frx":14C1A
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
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
         Height          =   700
         Left            =   480
         Picture         =   "INIGES04.frx":16FB2
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Ordenes de Trabajo y Fabricación"
         Top             =   3710
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Presup"
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
         Left            =   180
         Picture         =   "INIGES04.frx":172BC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Sub-Sistema de Presupuestación de Cámaras"
         Top             =   4420
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Repara"
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
         Left            =   180
         Picture         =   "INIGES04.frx":175C6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Sub-Sistema de Reparaciones"
         Top             =   3710
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Direct"
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
         Left            =   180
         Picture         =   "INIGES04.frx":178D0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Accesos Directos a Otras Aplicaciones"
         Top             =   5130
         Width           =   700
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Reset"
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
         Left            =   180
         Picture         =   "INIGES04.frx":17BDA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   1650
         Width           =   700
      End
      Begin VB.CommandButton Command31 
         Caption         =   "M.O.D"
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
         Left            =   180
         Picture         =   "INIGES04.frx":17D24
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Sub-Sistema de Control de Mano de Obra"
         Top             =   4470
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stock"
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
         Left            =   180
         Picture         =   "INIGES04.frx":1802E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sistema de Gestion de Stocks"
         Top             =   3000
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Util."
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
         Left            =   180
         Picture         =   "INIGES04.frx":18338
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Acceso a Utilitarios"
         Top             =   7080
         Width           =   700
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Setup"
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
         Left            =   180
         Picture         =   "INIGES04.frx":18642
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   6360
         Width           =   700
      End
      Begin VB.CommandButton Command2 
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
         Left            =   180
         Picture         =   "INIGES04.frx":18A84
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda FLEXOFT en Linea"
         Top             =   950
         Width           =   700
      End
      Begin VB.CommandButton Command1 
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
         Left            =   180
         Picture         =   "INIGES04.frx":18D8E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   240
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7590
      Picture         =   "INIGES04.frx":18ED8
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   720
      Width           =   2115
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "    Tesorería   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   420
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7245
      Picture         =   "INIGES04.frx":2089E
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   5415
      Width           =   2010
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Acreedores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   350
         TabIndex        =   21
         Top             =   460
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1365
      Picture         =   "INIGES04.frx":28264
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   5340
      Width           =   2115
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   630
         TabIndex        =   20
         Top             =   465
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   945
      Picture         =   "INIGES04.frx":2FD5E
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   645
      Width           =   2010
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivos Maestros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   420
         TabIndex        =   23
         Top             =   300
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   2925
      Picture         =   "INIGES04.frx":379F0
      ScaleHeight     =   1485
      ScaleWidth      =   4845
      TabIndex        =   4
      Top             =   3000
      Width           =   4845
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11925
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line7 
      X1              =   5250
      X2              =   8190
      Y1              =   3720
      Y2              =   1725
   End
   Begin VB.Line Line6 
      X1              =   5250
      X2              =   7980
      Y1              =   3750
      Y2              =   5640
   End
   Begin VB.Line Line4 
      X1              =   2925
      X2              =   5235
      Y1              =   5535
      Y2              =   3750
   End
   Begin VB.Line Line1 
      X1              =   1950
      X2              =   5310
      Y1              =   1425
      Y2              =   3735
   End
   Begin VB.Menu Archima 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Menuarch 
         Caption         =   "Menu Archivos Maestros"
      End
      Begin VB.Menu plh5 
         Caption         =   "-"
      End
      Begin VB.Menu Amasto 
         Caption         =   "Items de Stock"
      End
      Begin VB.Menu plc1 
         Caption         =   "-"
      End
      Begin VB.Menu Amacli 
         Caption         =   "Maestro de Clientes "
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu Amapro 
         Caption         =   "Proveedores y Gastos"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Ven_Com 
      Caption         =   "Ventas"
      Begin VB.Menu Pre_clientes 
         Caption         =   "Sub-Sistema de Emisión de Presupuestos"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh14 
         Caption         =   "-"
      End
      Begin VB.Menu ped_clientes 
         Caption         =   "Sub-Sistema de Pedidos de Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu PLH151 
         Caption         =   "-"
      End
      Begin VB.Menu Remitos 
         Caption         =   "Gestión de Remitos"
         Enabled         =   0   'False
      End
      Begin VB.Menu PLH15 
         Caption         =   "-"
      End
      Begin VB.Menu Ventas 
         Caption         =   "Gestión Ventas y Deudores"
         Begin VB.Menu Fac_ven 
            Caption         =   "Facturacion en Cuenta Corriente"
            Enabled         =   0   'False
         End
         Begin VB.Menu Recibo1 
            Caption         =   "Recibo de Cobranza"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu plh13 
         Caption         =   "-"
      End
      Begin VB.Menu Lis_pre 
         Caption         =   "Listas de Precios de Venta"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Acredor 
      Caption         =   "Compras y Proveedores"
      Begin VB.Menu compras 
         Caption         =   "Sub-Sistema de Compra de Mercaderias"
         Enabled         =   0   'False
      End
      Begin VB.Menu Fac_provee 
         Caption         =   "Facturacion de Compras"
         Enabled         =   0   'False
      End
      Begin VB.Menu Orpago1 
         Caption         =   "Ordenes de Pago"
         Enabled         =   0   'False
      End
      Begin VB.Menu subdiarioyestadisticas 
         Caption         =   "Subdiarios-Estadísticas de Compras"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Tesor 
      Caption         =   "Tesoreria"
      Begin VB.Menu Cob_Pag 
         Caption         =   "Cobranzas y Pagos"
         Begin VB.Menu Recibo 
            Caption         =   "Recibo de Cobranza"
            Enabled         =   0   'False
         End
         Begin VB.Menu Orpago 
            Caption         =   "Ordenes de Pago"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu plh17 
         Caption         =   "-"
      End
      Begin VB.Menu accesoamodulotesoreria 
         Caption         =   "Acceso Directo a Módulo de Tesorería"
      End
      Begin VB.Menu repara 
         Caption         =   "Sub-Sistema de Reparaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu CONMOD 
         Caption         =   "Sub-Sistema Control Mano de Obra"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu especificos 
      Caption         =   "Específicos"
      Begin VB.Menu recepcion 
         Caption         =   "Recepción de Mercadería"
         Enabled         =   0   'False
      End
      Begin VB.Menu ctrl_stock 
         Caption         =   "Control y Gestión de Stocks"
         Enabled         =   0   'False
      End
      Begin VB.Menu producción 
         Caption         =   "Producción"
         Enabled         =   0   'False
      End
      Begin VB.Menu MRP 
         Caption         =   "MRP - Planificación de los requerimientos de material"
         Enabled         =   0   'False
      End
      Begin VB.Menu contab_gen 
         Caption         =   "Sistema de Contabilidad General"
         Enabled         =   0   'False
      End
      Begin VB.Menu Anacos 
         Caption         =   "Analisis de Costos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu config 
      Caption         =   "Configuración"
      Begin VB.Menu configgeneral 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INIGES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub accesoamodulotesoreria_Click()
     GESVAL04.Show 1
End Sub

Private Sub Amacli_Click()
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     ACONEC$ = "ARCHIG04/AMACLI"
     Call CONECRUN(ACONEC$, "Maestro de Clientes y Deudores")
   Else
     Call COMUNI("Falta Módulo ARCHIG04")
   End If

'   If EXISTE%("AMACLI02.EXE") > 0 Then
'      Call CONECRUN("AMACLI02", "Maestro de Clientes y Deudores")
'    Else
'      Call CONECRUN("ARCH_GES/AMACLI", "Maestro de Items de Stock")
'   End If
End Sub

Private Sub Amapro_Click()
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     ACONEC$ = "ARCHIG04/AMAPRO"
     Call CONECRUN(ACONEC$, "Maestro de Proveedores y Sub-Contratistas")
   Else
     Call COMUNI("Falta Módulo ARCHIG04")
   End If
'   Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
End Sub

Private Sub Amasto_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If EXISTE%("AMASTO02.EXE") > 0 Then
       Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
     Else
       Call CONECRUN("ARCHIG04/AMASTO", "Maestro de Items de Stock")
   End If
End Sub

Private Sub Bus_che_Click()
  GESVAL04.Option6.Value = True
  Call GESVAL04.Command7_Click
'  GESVAL00.Option6.Value = True
'  Call GESVAL00.Command7_Click
End Sub

Private Sub bus_che1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Busqueda de Cheque
  Call Bus_che_Click
End Sub

Private Sub cgc_acree_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES04/CGCOMPRAS"
     Call CONECRUN(ACONEC$, "Control Gestión de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES04")
   End If
'   Call CONECRUN("ACRE_GES/CGCOMPRAS", "Control Gestion de Compras")
End Sub

Private Sub cgc_com1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/
  ' Control Gestion de Compras
  Call cgc_acree_Click
End Sub

Private Sub cgc_ven_Click()
   Call CONECRUN("VENGES04/CGVENTAS", "Control Gestion de Ventas")
End Sub

Private Sub cgc_ven1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/
  ' Control Gestion de Ventas
  Call cgc_ven_Click
End Sub

Private Sub Anacos_Click()
   If EXISTE%("ANACOS04.EXE") > 0 Then
     ACONEC$ = "ANACOS04"
     Call CONECRUN(ACONEC$, "Analisis de Costos")
   Else
     Call COMUNI("Falta Módulo ANACOS04")
   End If

End Sub

Private Sub Command1_Click()
  If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
     Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  End If
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  If PROCACTI% < 1 Then
    Call GENBACKUP
    Call CANCELUCOM
    Call FINAL("")
  End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call CONECRUN("ORCOMP04", "Sub-Sistema de Compras")
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   'ARCHIM00.Show 1
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("ARCHIG04", "A-B-M-L Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("VENGES04", "Gestión Ventas y Deudores")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If DERACCE%("GESTVAL", "Gestion de Cheques y Valores") >= 2 Then
     GESVAL04.Show 1
   End If
   Command13.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("ACRGES04", "Gestión Compras y Acreedores")
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
     Call CANCELUCOM
99   ' conecta con FLOGON.EXE
     On Error Resume Next
     RETOR$ = TRIM$(UCase$(App.EXEName))
     Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
     If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
     End If
     Call FINAL("")
   End If
   '
   Command16.Enabled = True
End Sub

Private Sub Command17_Click()

End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("CONTAB04.EXE") > 0 Then
       Call CONECRUN("CONTAB04", "Contabilidad General")
     ElseIf EXISTE%("CONTAGEN.EXE") > 0 Then
       Call SELEJER(EJUI%)
       If Abs(EJUI%) <> 1 Then GoTo 99
       Call CONECRUN("CONTAGEN", "Contabilidad General")
   End If
99 Command18.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO04"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   TTXX$ = ""
   Dim ACONE$(16), DECONE$(16)
   ICONE% = 0
   If EXISTE%("FLGEST04.EXE") > 0 Then
      TTXX$ = TTXX$ + "FLEXOFT Industrial\"
      ICONE% = ICONE% + 1: ACONE$(ICONE%) = "FLEXO04": DECONE$(ICONE%) = "Sistema de Gestión Industrial"
   End If
   If EXISTE%("APROYE04.EXE") > 0 Then
     TTXX$ = TTXX$ + "Administacion de Proyectos\"
     ICONE% = ICONE% + 1: ACONE$(ICONE%) = "APROYE04": DECONE$(ICONE%) = "Administración de Proyectos"
   End If
   '
   If TTXX$ = "" Then GoTo 5
   '
   TTXX$ = TTXX$ + "Otros Accesos Directos"
   OPX% = COMALTER%("Opciones de Conexión a Accesos Directos:\\" + TTXX$)
   If OPX% > 0 Then
      If OPX% <= ICONE% Then
           Call CONECRUN(ACONE$(OPX%), DECONE$(OPX%))
           GoTo 99
         Else
           GoTo 5
      End If
   End If
   GoTo 99
   '
5  Call ACCDIR("FLGEST04")
   '
99 Command3.Enabled = True
   '
End Sub

Private Sub Command31_Click()
   Command31.Enabled = False
   Call CONECRUN("CONMOD03", "Control de Mano de Obra")
   Command31.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   FSETUPNEW.Show 1
'   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If EXISTE%("CONSTO04.EXE") > 0 Then
       Call CONECRUN("CONSTO04", "Control Básico de Stocks")
     ElseIf EXISTE%("STOBAS04.EXE") > 0 Then
       Call CONECRUN("STOBAS04", "Control Básico de Stocks")
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Version")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If EXISTE%("PEDIDO04.EXE") > 0 And EXISTE%("PRESUP04.EXE") > 0 Then
      OPX% = ALTERNA%("Emisión de Presupuestos\Pedidos de Clientes")
      If OPX% = 1 Then GoTo 10
      If OPX% = 2 Then GoTo 20
      GoTo 99
   End If
   '
   If EXISTE%("PRESUP04.EXE") > 0 Then
10    Call CONECRUN("PRESUP04", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
   If EXISTE%("PEDIDO04.EXE") > 0 Then
20    Call CONECRUN("PEDIDO04", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("REPARA02", "Sub-Sistema de Reparaciones")
   Command8.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
   Command14.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("PREFRI04.EXE") > 0 Then
        Call CONECRUN("PREFRI04", "Presupuestación Cámaras Frigoríficas")
      ElseIf EXISTE%("PREFLE04.EXE") > 0 Then
        Call CONECRUN("PREFLE04", "Presupuestación Sistemas FLEXOFT")
      ElseIf EXISTE%("ANACOS04.EXE") > 0 Then
        Call CONECRUN("ANACOS04", "Análisis de Costos Standard")
      ElseIf EXISTE%("ANACOS06.EXE") > 0 Then
        Call CONECRUN("ANACOS06", "Análisis de Costos Standard")
   End If
   Command9.Enabled = True
End Sub

Private Sub compras_Click()
   Call CONECRUN("ORCOMP04", "Sub-Sistema de Compras")
End Sub

Private Sub compras1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Proveedores/Sub-Sistema de Compras
  Call compras_Click
End Sub

Private Sub configgeneral_Click()
    Call Command4_Click

End Sub

Private Sub CONMOD_Click()
   Call Command31_Click

End Sub

Private Sub contab_gen_Click()
   Call CONECRUN("CONTAB04", "Contabilidad General")
End Sub

Private Sub ctrl_stock_Click()
   Call CONECRUN("CONSTO04", "Control de Stocks")
End Sub

Private Sub Deposito_Click()
   Call GESVAL04.Command11_Click
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Fac_provee_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES04/FACTURA"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES04")
   End If
'   Call CONECRUN("ACRE_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Fac_ven_Click()
   Call CONECRUN("VENGES04/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Form_Load()
   '
   BCOL% = 192
   BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture1.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture2.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture3.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture4.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture6.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   Picture8.BackColor = RGB(BCOL%, BCOL%, BCOL%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture5.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
    End If

   EJER% = 1
   Call GRARAM(69, Chr$(1))
   '
   If EXISTE%("CONTAB04.EXE") > 0 Then
      Command18.Visible = True
      contab_gen.Enabled = True
   End If
   If EXISTE%("CONSTO04.EXE") > 0 Or EXISTE%("STOBAS04.EXE") > 0 Then
      Command5.Visible = True
      ctrl_stock.Enabled = True
   End If
   If EXISTE%("ORFABR04.EXE") > 0 Then
       Command14.Visible = True
       Command14.Left = Command8.Left
       producción.Enabled = True
     ElseIf EXISTE%("REPARA02.EXE") > 0 Then
       Command8.Visible = True
       repara.Enabled = True
   End If
   If EXISTE%("CONMOD04.EXE") > 0 Then
     Command31.Visible = True
     CONMOD.Enabled = True
   End If
   If EXISTE%("PEDIDO04.EXE") > 0 Then
      Command7.Visible = True
      Command7.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
      ped_clientes.Enabled = True
   End If
   If EXISTE%("PRESUP04.EXE") > 0 Then
      Command7.ToolTipText = "Sub-Sistema de Presupuestos a de Clientes"
      Command7.Visible = True
      Pre_clientes.Enabled = True
   End If
   If EXISTE%("PEDIDO04.EXE") > 0 And EXISTE%("PRESUP04.EXE") > 0 Then
      Command7.ToolTipText = "Sub-Sistemas de Presupuestos y Pedidos de Clientes"
   End If
   If EXISTE%("REMITO04.EXE") > 0 Then
     Command23.Visible = True
   End If
   '
   If EXISTE%("ORCOMP04.EXE") > 0 Then
      Command10.Visible = True
      compras.Enabled = True
   End If
   '
   
   If EXISTE%("ANACOS04.EXE") > 0 Then
     Command9.ToolTipText = "Módulo de Análisis de Costos de Productos"
     Command9.Caption = "An-Cos"
     Command9.Visible = True
   End If
   '
   If EXISTE%("ANACOS06.EXE") > 0 Then
     Command9.ToolTipText = "Módulo de Análisis de Costos de Productos"
     Command9.Caption = "An-Cos"
     Command9.Visible = True
   End If
   '
   If EXISTE%("PREFRI04.EXE") > 0 Or EXISTE%("PREFLE04.exe") > 0 Then
     Command9.Visible = True
   End If
   '
   If EXISTE%("ARCHIG04.EXE") > 0 Then
      Amapro.Enabled = True
   End If
   '
   If EXISTE%("REMITO04.EXE") > 0 Then
         Remitos.Enabled = True
   End If
   
   If EXISTE%("VENGES04.EXE") > 0 Then
         Fac_ven.Enabled = True
         Recibo1.Enabled = True
         Recibo.Enabled = True
   End If
   '
   If EXISTE%("LISPRE04.EXE") > 0 Then
         Lis_pre.Enabled = True
   End If
   
   If EXISTE%("ORCOMP04.EXE") > 0 Then
         compras.Enabled = True
   End If
   
   If EXISTE%("ACRGES04.EXE") > 0 Then
         Fac_provee.Enabled = True
         Orpago1.Enabled = True
         subdiarioyestadisticas.Enabled = True
         Orpago.Enabled = True
   End If
   '
   If EXISTE%("BRECEP04.EXE") > 0 Then
         recepcion.Enabled = True
   End If
   If EXISTE%("PROMRP04.EXE") > 0 Then
         MRP.Enabled = True
   End If
   '
   If EXISTE%("ANACOS04.EXE") > 0 Then
         Anacos.Enabled = True
   End If
   
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
   '
   
End Sub

Private Sub Form_Resize()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
        Picture5.Width = Me.Width - (Picture16.Width + 100)
        Picture5.Height = Me.Height - 50
        Image2.Width = Me.Width - (Picture16.Width + 100)
        Image2.Height = Me.Height - 100
        Picture7.Left = Picture5.Width
        Picture7.Height = Me.Height - 50
        Picture16.Height = Picture7.Height
        Image3.Height = Me.Height - 100
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
     Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   End If
   Call CANCELUCOM
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub help_Click()
  MsgBox "Sysflom Soluciones Informáticas" & vbCrLf & _
         "Teléfono: (00 54 11) 4862-0119 (rotativas)" & vbCrLf & _
         "Email: asistec@sysflom.com.ar", vbQuestion, _
         "Mesa de Ayuda "
End Sub

Private Sub Lis_pre_Click()
   If EXISTE%("LISPRE04.EXE") > 0 Then
     Call CONECRUN("LISPRE04", "Lista de Precios")
   Else
     Call COMUNI("Falta Módulo LISPRE04")
   End If
'   Call CONECRUN("VENGES04/LPRECIOS", "Lista de Precios")
End Sub

Private Sub Menuarch_Click()
   Call Command11_Click
End Sub

Private Sub MRP_Click()
   Call CONECRUN("PROMRP04", "MRP - Planificación de los requerimientos de material")
End Sub

Private Sub ordefab_Click()
    Call Command14_Click
End Sub

Private Sub Orpago_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES04/ORPAGO"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES04")
   End If
End Sub

Private Sub Orpago1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cobranzas y Pagos/Ordenes de Pago
  Call Orpago_Click
End Sub

Private Sub ped_clientes_Click()
  Call CONECRUN("PEDIDO04", "Pedidos de Clientes")
End Sub

Private Sub pre_clientes_Click()
  Call CONECRUN("PRESUP04", "Presupuestos a Clientes")
End Sub

Private Sub ped_clientes1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/Pedido de Clientes
  Call ped_clientes_Click
End Sub

Private Sub print1_Click()
  GESVAL04.Option2.Value = True ' Opcion diferidos
  GESVAL04.Option4.Value = True ' Opcion por impresora
  Call GESVAL04.Command2_Click ' Ejecuta la consulta
'  GESVAL00.Option2.Value = True ' Opcion diferidos
'  GESVAL00.Option4.Value = True ' Opcion por impresora
'  Call GESVAL00.Command2_Click ' Ejecuta la consulta
End Sub

Private Sub print12_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Propios Diferidos/Impresora
  Call print1_Click
End Sub

Private Sub print2_Click()
  GESVAL04.Option1.Value = True ' Opcion cartera
  GESVAL04.Option4.Value = True ' Opcion impresora
  Call GESVAL04.Command2_Click ' Ejecuta consulta
'  GESVAL00.Option1.Value = True ' Opcion cartera
'  GESVAL00.Option4.Value = True ' Opcion impresora
'  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub print22_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Cartera de Valores/Impresora
   Call print2_Click
End Sub

Private Sub producción_Click()
       Call CONECRUN("ORFABR04", "Sub-Sistema de Fabricación")
End Sub

Private Sub recepcion_Click()
   If EXISTE%("BRECEP04.EXE") > 0 Then
     ACONEC$ = "BRECEP04"
     Call CONECRUN(ACONEC$, "Recepción de Mercadería")
   Else
     Call COMUNI("Falta Módulo BRECEP04")
   End If

End Sub

Private Sub Recibo_Click()
   If EXISTE%("VENGES04.EXE") > 0 Then
     ACONEC$ = "VENGES04/RECIBO"
     Call CONECRUN(ACONEC$, "Recibo de Cobranza")
   Else
     Call COMUNI("Falta Módulo VENGES04")
   End If
End Sub

Private Sub Recibo1_Click()
   ' Realiza una llamada a la opcion de menu
   ' Tesoreria/Cobranzas y Pagos/Recibo de Cobranza
   Call Recibo_Click
End Sub

Private Sub Remitos_Click()
    Call Command23_Click
End Sub

Private Sub repara_Click()
   Call CONECRUN("REPARA02", "Sub-Sistema de Reparaciones")
End Sub

Private Sub Res_cuenta_Click()
'  GESVAL00.Option5.Value = True ' opcion impresora
'  Call GESVAL00.Command17_Click ' Ejecuta la consulta
End Sub

Private Sub res_cuenta1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Consultas e Informes\Resumen de Cuenta
  Call Res_cuenta_Click
End Sub

Private Sub Sal_Tes_Click()
  GESVAL04.Option8.Value = True ' Opcion saldos
  GESVAL04.Option5.Value = True ' Opcion pantalla
  Call GESVAL04.Command2_Click ' Ejecuta consulta
'  GESVAL00.Option8.Value = True ' Opcion saldos
'  GESVAL00.Option5.Value = True ' Opcion pantalla
'  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub Sal_Tes1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Saldos de Tesoreria
  Call Sal_Tes_Click
End Sub

Private Sub scr1_Click()
  GESVAL04.Option2.Value = True ' Opcion diferidos
  GESVAL04.Option5.Value = True ' Opcion pantalla
  Call GESVAL04.Command2_Click ' Ejecuta consulta
'  GESVAL00.Option2.Value = True ' Opcion diferidos
'  GESVAL00.Option5.Value = True ' Opcion pantalla
'  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub scr11_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Propios Diferidos/Pantalla
  Call scr1_Click
End Sub

Private Sub scr2_Click()
  GESVAL04.Option1.Value = True ' Opcion cartera
  GESVAL04.Option5.Value = True ' Opcion pantalla
  Call GESVAL04.Command2_Click ' Ejecuta consulta
'  GESVAL00.Option1.Value = True ' Opcion cartera
'  GESVAL00.Option5.Value = True ' Opcion pantalla
'  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub scr21_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Cartera de Valores/Pantalla
  Call scr2_Click
End Sub

Private Sub Subdia_CajaBan_Click()
'  GESVAL00.Option4.Value = True ' opcion impresora
'  Call GESVAL00.Command15_Click ' Ejecuta la consulta
End Sub

Private Sub Subdia_CajaBan1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Consultas e Informes\Subdiario de Caja y Bancos
  Call Subdia_CajaBan_Click
End Sub

Private Sub subdiarioyestadisticas_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES04/SUBDESTA"
     Call CONECRUN(ACONEC$, "Sub-Diario y Estadísticas de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES04")
   End If

End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
End Sub

Private Sub Trafondo_Click()
   Call GESVAL04.Command10_Click
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image1.Picture = LoadPicture("")
        On Error Resume Next
        Image1.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        ' Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
        If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") < 1 Then ' PARA CASO GENTILI
           Picture9.Visible = True
           IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
           '
           FORMALTER.Image3.Picture = LoadPicture("")
           On Error Resume Next
           FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
           On Error GoTo 0
           GoTo 12
        End If
        '
      End If
   End If
   '
   Picture9.Visible = False
   If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") > 0 Then
     Picture9.Visible = True
     IMAGEFILE$ = LUDAT$ + "LOGOEMPR.BMP"
     '
     FORMALTER.Image3.Picture = LoadPicture("")
     FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
     '
12   HIMAGEN = FORMALTER.Image3.Picture.Width
     VIMAGEN = FORMALTER.Image3.Picture.Height
     '
     HCUADRO = Picture9.Width
     VCUADRO = Picture9.Height
     '
     REDUH = HIMAGEN / HCUADRO
     REDUV = VIMAGEN / VCUADRO
     REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
     If REDUC < 1 Then REDUC = 1
     '
     With Picture9
       .Height = VCUADRO * REDUV / REDUC
       .Width = HCUADRO * REDUH / REDUC
       .Left = .Left + (HCUADRO - Picture9.Width) / 2
         If .Left < 300 Then Left = 300
       .Top = .Top - (VCUADRO - Picture9.Height) / 2
         If .Top < 400 Then .Top = 400
     End With
     '
     Image1.Width = Picture9.Width
     Image1.Height = Picture9.Height
     Image1.Stretch = True
     On Error Resume Next
     Image1.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
     On Error GoTo 0
     '
     FORMALTER.Image3.Picture = LoadPicture("")
   End If
   '
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      x$ = TRIM$(UCase$(File1.List(U& - 1)))
      If x$ <> "" Then
        If x$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + x$) > 0 Then
            If Left$(x$, 7) = "FLEXO04" Then GoTo 9
            If Left$(x$, 8) = "FLGEST04" Then GoTo 9
            If Left$(x$, 8) = "CONTAB04" Then GoTo 9
            On Error GoTo 10
            Kill LUCOM$ + x$
            On Error GoTo 0
          End If
        End If
      End If
9  Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(x$, Len(x$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function

