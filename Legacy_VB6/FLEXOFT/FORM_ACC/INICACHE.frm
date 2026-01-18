VERSION 5.00
Begin VB.Form INICACHE 
   BackColor       =   &H00C2C2C2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Gestión"
   ClientHeight    =   8040
   ClientLeft      =   1935
   ClientTop       =   2310
   ClientWidth     =   11880
   Icon            =   "INICACHE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9975
      Picture         =   "INICACHE.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Sistema de Contabilidad General"
      Top             =   2835
      Visible         =   0   'False
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
      Left            =   10020
      Picture         =   "INICACHE.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Informes de Producción"
      Top             =   5730
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command19 
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
      Left            =   8400
      Picture         =   "INICACHE.frx":2476
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Ordenes de Trabajo y Fabricación"
      Top             =   4725
      Width           =   600
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8505
      Picture         =   "INICACHE.frx":2780
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   34
      Top             =   4935
      Width           =   2010
   End
   Begin VB.PictureBox Picture11 
      Height          =   2115
      Left            =   525
      ScaleHeight     =   2055
      ScaleWidth      =   1950
      TabIndex        =   31
      Top             =   525
      Visible         =   0   'False
      Width           =   2010
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   1905
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
      Height          =   600
      Left            =   8400
      Picture         =   "INICACHE.frx":A54A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Tesorería y Gestión de Valores"
      Top             =   1785
      Width           =   600
   End
   Begin VB.CommandButton Command17 
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
      Left            =   2985
      Picture         =   "INICACHE.frx":A854
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5565
      Width           =   600
   End
   Begin VB.CommandButton Command16 
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
      Left            =   5880
      Picture         =   "INICACHE.frx":AB5E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5880
      Width           =   600
   End
   Begin VB.PictureBox Picture10 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3150
      Picture         =   "INICACHE.frx":AE68
      ScaleHeight     =   1140
      ScaleWidth      =   2010
      TabIndex        =   29
      Top             =   5800
      Width           =   2010
   End
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   1150
      Left            =   6105
      Picture         =   "INICACHE.frx":12A96
      ScaleHeight     =   1155
      ScaleWidth      =   2010
      TabIndex        =   28
      Top             =   6120
      Width           =   2010
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1150
      Left            =   8505
      Picture         =   "INICACHE.frx":1A590
      ScaleHeight     =   1155
      ScaleWidth      =   2010
      TabIndex        =   24
      Top             =   1980
      Width           =   2010
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tesorería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   190
         TabIndex        =   25
         Top             =   465
         Width           =   1725
      End
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
      Height          =   600
      Left            =   4920
      Picture         =   "INICACHE.frx":21F56
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Generación de Presupuestos a Clientes"
      Top             =   1110
      Visible         =   0   'False
      Width           =   600
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
      Left            =   7245
      Picture         =   "INICACHE.frx":22260
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sub-Sistema de Compras Industriales"
      Top             =   1515
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
      Left            =   4440
      Picture         =   "INICACHE.frx":2256A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sub-Sistema de Pedidos de Clientes"
      Top             =   1830
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
      Left            =   5670
      Picture         =   "INICACHE.frx":229AC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Acreedores - Cuentas por Pagar"
      Top             =   420
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
      Left            =   2880
      Picture         =   "INICACHE.frx":22DEE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Gestión de Ventas y Deudores"
      Top             =   735
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
      Left            =   195
      Picture         =   "INICACHE.frx":230F8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   2970
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1250
      Left            =   210
      Picture         =   "INICACHE.frx":23402
      ScaleHeight     =   1245
      ScaleWidth      =   1800
      TabIndex        =   5
      Top             =   6615
      Width           =   1800
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   10800
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command5 
         Caption         =   "Costos"
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
         Picture         =   "INICACHE.frx":2AB94
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Costos Industriales - Presupuestación de Productos"
         Top             =   4095
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
         Picture         =   "INICACHE.frx":2AE9E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Acceso a Utilitarios"
         Top             =   7080
         Width           =   700
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Tablas"
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
         Picture         =   "INICACHE.frx":2B1A8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Actualizacion de Tablas de Validacion"
         Top             =   5640
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
         Picture         =   "INICACHE.frx":2B4B2
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "INICACHE.frx":2B8F4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda FLEXOFT en Linea"
         Top             =   960
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
         Picture         =   "INICACHE.frx":2BBFE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   240
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1150
      Left            =   5900
      Picture         =   "INICACHE.frx":2BD48
      ScaleHeight     =   1155
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   615
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
         TabIndex        =   19
         Top             =   465
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3045
      Picture         =   "INICACHE.frx":3370E
      ScaleHeight     =   1140
      ScaleWidth      =   2010
      TabIndex        =   1
      Top             =   945
      Width           =   2010
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
         TabIndex        =   18
         Top             =   450
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   420
      Picture         =   "INICACHE.frx":3B208
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   3165
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
         TabIndex        =   20
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
      Picture         =   "INICACHE.frx":42E9A
      ScaleHeight     =   1485
      ScaleWidth      =   4845
      TabIndex        =   3
      Top             =   3000
      Width           =   4845
   End
   Begin VB.Image Image3 
      Height          =   1590
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Line Line8 
      X1              =   5250
      X2              =   7245
      Y1              =   3885
      Y2              =   6720
   End
   Begin VB.Line Line5 
      X1              =   5250
      X2              =   4095
      Y1              =   3780
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   4935
      X2              =   6720
      Y1              =   3885
      Y2              =   1470
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11925
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line7 
      X1              =   5250
      X2              =   9555
      Y1              =   3720
      Y2              =   2415
   End
   Begin VB.Line Line6 
      X1              =   5250
      X2              =   9450
      Y1              =   3750
      Y2              =   5250
   End
   Begin VB.Line Line4 
      X1              =   4095
      X2              =   5235
      Y1              =   6405
      Y2              =   3750
   End
   Begin VB.Line Line1 
      X1              =   2100
      X2              =   5310
      Y1              =   3750
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
      Begin VB.Menu Amapro 
         Caption         =   "Proveedores y Gastos"
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Ven_Com 
      Caption         =   "Aplicaciones Comerciales"
      Begin VB.Menu Ventas 
         Caption         =   "Gestión Ventas y Deudores"
         Begin VB.Menu Fac_ven 
            Caption         =   "Facturacion en Cuenta Corriente"
         End
         Begin VB.Menu Recibo1 
            Caption         =   "Recibo de Cobranza"
         End
         Begin VB.Menu plh13 
            Caption         =   "-"
         End
         Begin VB.Menu Lis_pre 
            Caption         =   "Listas de Precios de Venta"
         End
         Begin VB.Menu plh14 
            Caption         =   "-"
         End
         Begin VB.Menu cgc_ven 
            Caption         =   "Control de Gestion de Ventas"
         End
      End
      Begin VB.Menu plh26 
         Caption         =   "-"
      End
      Begin VB.Menu ped_clientes 
         Caption         =   "Sub-Sistema de Pedidos de Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu pres_clientes 
         Caption         =   "Sub-Sistema de Emision de Presupuestos"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh15 
         Caption         =   "-"
      End
      Begin VB.Menu Acredor 
         Caption         =   "Compras y Proveedores"
         Begin VB.Menu Fac_provee 
            Caption         =   "Facturacion de Compras"
         End
         Begin VB.Menu Orpago1 
            Caption         =   "Ordenes de Pago"
         End
         Begin VB.Menu plh16 
            Caption         =   "-"
         End
         Begin VB.Menu cgc_acree 
            Caption         =   "Control Gestion de Compras"
         End
      End
      Begin VB.Menu compras 
         Caption         =   "Sub-Sistema de Compras Industriales"
      End
   End
   Begin VB.Menu Tesor 
      Caption         =   "Tesoreria"
      Begin VB.Menu Cob_Pag 
         Caption         =   "Cobranzas y Pagos"
         Begin VB.Menu Recibo 
            Caption         =   "Recibo de Cobranza"
         End
         Begin VB.Menu Orpago 
            Caption         =   "Ordenes de Pago"
         End
      End
      Begin VB.Menu plh17 
         Caption         =   "-"
      End
      Begin VB.Menu Mov_Banc 
         Caption         =   "Movimientos Bancarios"
         Begin VB.Menu Deposito 
            Caption         =   "Depósito Bancario"
         End
         Begin VB.Menu Trafondo 
            Caption         =   "Transferencia de Fondos"
         End
      End
      Begin VB.Menu plh18 
         Caption         =   "-"
      End
      Begin VB.Menu Che_Car 
         Caption         =   "Cheques y Valores"
         Begin VB.Menu checar 
            Caption         =   "Listado de Valores en Cartera"
            Begin VB.Menu scr2 
               Caption         =   "Pantalla"
            End
            Begin VB.Menu print2 
               Caption         =   "Impresora"
            End
         End
         Begin VB.Menu che_dif 
            Caption         =   "Cheques Propios Diferidos"
            Begin VB.Menu scr1 
               Caption         =   "Pantalla"
            End
            Begin VB.Menu print1 
               Caption         =   "Impresora"
            End
         End
         Begin VB.Menu Sal_Tes 
            Caption         =   "Listado de Saldos de Tesoreria"
         End
         Begin VB.Menu plh3 
            Caption         =   "-"
         End
         Begin VB.Menu bus_che 
            Caption         =   "Busqueda de Cheque por Número"
         End
      End
      Begin VB.Menu Icontab 
         Caption         =   "Consultas e Informes"
         Begin VB.Menu Subdia_CajaBan 
            Caption         =   "Subdiario de Caja y Bancos"
         End
         Begin VB.Menu res_cuenta 
            Caption         =   "Resumen de Cuenta"
         End
      End
   End
   Begin VB.Menu cons_inf 
      Caption         =   "Consultas e Informes"
      Begin VB.Menu checar1 
         Caption         =   "Cartera de Valores"
         Begin VB.Menu scr21 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu print22 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu che_dif1 
         Caption         =   "Cheques Propios Diferidos"
         Begin VB.Menu scr11 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu print12 
            Caption         =   "Imprimir / Exportar"
         End
      End
      Begin VB.Menu Sal_Tes1 
         Caption         =   "Saldos de Tesoreria"
      End
      Begin VB.Menu plh7 
         Caption         =   "-"
      End
      Begin VB.Menu bus_che1 
         Caption         =   "Busqueda de Cheque"
      End
      Begin VB.Menu plh12 
         Caption         =   "-"
      End
      Begin VB.Menu Subdia_CajaBan1 
         Caption         =   "Subdiario de Caja y Bancos"
      End
      Begin VB.Menu res_cuenta1 
         Caption         =   "Resumen de Cuenta C & B"
      End
      Begin VB.Menu plh9 
         Caption         =   "-"
      End
      Begin VB.Menu cgc_ven1 
         Caption         =   "Control de Gestion de Ventas"
      End
      Begin VB.Menu cgc_com1 
         Caption         =   "Control de Gestion de Compras"
      End
   End
   Begin VB.Menu comple 
      Caption         =   "Complementos"
      Begin VB.Menu ped_clientes1 
         Caption         =   "Sub-Sistema de Pedidos de Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu pres_clientes1 
         Caption         =   "Sub-Sistema de Emision de Presupuestos"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh27 
         Caption         =   "-"
      End
      Begin VB.Menu contab_gen 
         Caption         =   "Sistema de Contabilidad General"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh28 
         Caption         =   "-"
      End
      Begin VB.Menu ctrl_stock 
         Caption         =   "Control y Gestión de Stocks"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh8 
         Caption         =   "-"
      End
      Begin VB.Menu repara 
         Caption         =   "Sub-Sistema de Reparaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu CONMOD 
         Caption         =   "Sub-Sistema Control Mano de Obra"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh20 
         Caption         =   "-"
      End
      Begin VB.Menu compras1 
         Caption         =   "Sub-Sistema de Compras Industriales"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INICACHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Amacli_Click()
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     ACONEC$ = "ARCHIG04/AMACLI"
     Call CONECRUN(ACONEC$, "Maestro de Clientes y Deudores")
   Else
     Call COMUNI("Falta Módulo ARCHIG04")
   End If
   'If EXISTE%("AMACLI02.EXE") > 0 Then
   '    Call CONECRUN("AMACLI02", "Maestro de Clientes y Deudores")
   '  Else
   '    Call CONECRUN("ARCH_GES/AMACLI", "Maestro de Items de Stock")
   'End If
End Sub

Private Sub Amapro_Click()
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     ACONEC$ = "ARCHIG04/AMAPRO"
     Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Else
     Call COMUNI("Falta Módulo ARCHIG04")
   End If
   'Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
End Sub

Private Sub Amasto_Click()
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     ACONEC$ = "ARCHIG04/AMASTO"
     Call CONECRUN(ACONEC$, "Maestro de Items de Stock")
   Else
     Call COMUNI("Falta Módulo ARCHIG04")
   End If
   'Call CIERRARCH("*.*")
   'Call BAJALDISCO
   'If EXISTE%("AMASTO02.EXE") > 0 Then
   '    Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
   '  Else
   '    Call CONECRUN("ARCH_GES/AMASTO", "Maestro de Items de Stock")
   'End If
End Sub

Private Sub Bus_che_Click()
  GESVAL00.Option6.Value = True
  Call GESVAL00.Command7_Click
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
   'Call CONECRUN("ACRE_GES/CGCOMPRAS", "Control Gestion de Compras")
End Sub

Private Sub cgc_com1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/
  ' Control Gestion de Compras
  Call cgc_acree_Click
End Sub

Private Sub cgc_ven_Click()
   If EXISTE%("VENGES04.EXE") > 0 Then
     ACONEC$ = "VENGES04/CGVENTAS"
     Call CONECRUN("VENGES04", "Control Gestión de Ventas")
   Else
     Call COMUNI("Falta módulo VENGES04")
   End If
   'Call CONECRUN("VENT_GES/CGVENTAS", "Control Gestion de Ventas")
End Sub

Private Sub cgc_ven1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/
  ' Control Gestion de Ventas
  Call cgc_ven_Click
End Sub

Private Sub Command1_Click()
  If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
     Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  End If
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  Call GENBACKUP
  Call FINAL("")
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call CONECRUN("ORCOMP03", "Sub-Sistema de Compras")
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG04.EXE") > 0 Then
     Call CONECRUN("ARCHIG04", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG04")
   End If
   'ARCHIM00.Show 1
   'Call CIERRARCH("*.*")
   'Call BAJALDISCO
   'Call CONECRUN("ARCH_GES", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("VENGES04.EXE") > 0 Then
     Call CONECRUN("VENGES04", "Gestión Ventas y Deudores")
   Else
     Call COMUNI("Falta módulo VENGES04")
   End If
   'Call CIERRARCH("*.*")
   'Call BAJALDISCO
   'Call CONECRUN("VENT_GES", "Gestión Ventas y Deudores")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("IFABRI04.EXE") > 0 Then
     Call CONECRUN("IFABRI04", "Informes de Fabricación")
   Else
     Call COMUNI("Falta módulo IFABRI04")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If EXISTE%("CONTAB04.EXE") > 0 Then
     Call CONECRUN("CONTAB04", "Gestión de Cheques y Valores")
   Else
     Call COMUNI("Falta módulo CONTAB04")
   End If
   Command14.Enabled = True
End Sub

Sub Command15_Click()
   Command15.Enabled = False
   If EXISTE%("ACRGES04.EXE") > 0 Then
     Call CONECRUN("ACRGES04", "Gestión Compras y Acreedores")
   Else
     Call COMUNI("Falta módulo ACRGES04")
   End If
   Command15.Enabled = True
End Sub


Private Sub Command16_Click()
   Command16.Enabled = False
   If EXISTE%("ORCOMP04.EXE") > 0 Then
     Call CONECRUN("ORCOMP04", "Sistema de Ordenes de Compra")
   Else
     Call COMUNI("Falta módulo ORCOMP04")
   End If
   Command16.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   If EXISTE%("CONSTO04.EXE") > 0 Then
     Call CONECRUN("CONSTO04", "Sistema de Gestión de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO04")
   End If
   Command17.Enabled = True
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

Private Sub Command31_Click()
   Command31.Enabled = False
   Call CONECRUN("CONMOD03", "Control de Mano de Obra")
   Command31.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   If EXISTE%("ORFABR04.EXE") > 0 Then
     Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
   Else
     Call COMUNI("Falta módulo ORFABR04")
   End If
   Command19.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("ARCHIG04/TABLES", "Maestro de Articulos de Stock")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CONECRUN("ANACOS04", "Sistema de Analisis de Costos")
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Version")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("PEDIDO02", "Gestión de Pedidos de Clientes")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("REPARA02", "Sub-Sistema de Reparaciones")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("PRESUP02", "Gestión de Presupuestos a Clientes")
   Command9.Enabled = True
End Sub

Private Sub compras_Click()
   If EXISTE%("ORCOMP04.EXE") > 0 Then
     Call CONECRUN("ORCOMP04", "Sub-Sistema de Compras")
   Else
     Call COMUNI("Falta Módulo ORCOMP04")
   End If
   'Call CONECRUN("ORCOMP03", "Sub-Sistema de Compras")
End Sub

Private Sub compras1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Proveedores/Sub-Sistema de Compras
  Call compras_Click
End Sub

Private Sub CONMOD_Click()
   Call Command31_Click
End Sub

Private Sub contab_gen_Click()
   Call CONECRUN("CONTAB04", "Contabilidad General")
End Sub

Private Sub ctrl_stock_Click()
   Call CONECRUN("CONSTO02", "Control de Stocks")
End Sub

Private Sub Deposito_Click()
   Call MENUTES.Command8_Click
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
   'Call CONECRUN("ACRE_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Fac_ven_Click()
   If EXISTE%("VENGES04.EXE") > 0 Then
     ACONEC$ = "VENGES04/FACTURA"
     Call CONECRUN(ACONEC$, "Facturación de Ventas")
   Else
     Call COMUNI("Falta Módulo VENGES04")
   End If
   'Call CONECRUN("VENT_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Form_Load()
   '
   BCOL% = 193
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
   EJER% = 1
   Call GRARAM(69, Chr$(1))
   '
   If EXISTE%("CONTAB04.EXE") > 0 Then
      Command18.Visible = True
      contab_gen.Enabled = True
   End If
   If EXISTE%("CONSTO02.EXE") > 0 Then
      Command5.Visible = True
      ctrl_stock.Enabled = True
   End If
   If EXISTE%("REPARA02.EXE") > 0 Then
     Command8.Visible = True
     repara.Enabled = True
   End If
   If EXISTE%("CONMOD03.EXE") > 0 Then
     Command31.Visible = True
     CONMOD.Enabled = True
   End If
   If EXISTE%("PEDIDO02.EXE") > 0 Then
      Command7.Visible = True
      ped_clientes.Enabled = True
      ped_clientes1.Enabled = True
   End If
   If EXISTE%("PRESUP02.EXE") > 0 Then
      Command9.Visible = True
      pres_clientes.Enabled = True
      pres_clientes1.Enabled = True
   End If
   If EXISTE%("ORCOMP03.EXE") > 0 Then
      Command10.Visible = True
      compras.Enabled = True
      compras1.Enabled = True
   End If
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
     Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   End If
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub help_Click()
  MsgBox "FLEXOFT Soluciones Informáticas" & vbCrLf & _
         "Telefax: (00 54 11) 4541-1515 / 4524-1284" & vbCrLf & _
         "Email: asistec@flexoft.com.ar", vbQuestion, _
         "Mesa de Ayuda FLEXOFT"
End Sub

Private Sub Lis_pre_Click()
   If EXISTE%("LISPRE04.EXE") > 0 Then
     Call CONECRUN("LISPRE04", "Lista de Precios")
   Else
     Call COMUNI("Falta Módulo LISPRE04")
   End If
   'Call CONECRUN("VENT_GES/LPRECIOS", "Lista de Precios")
End Sub

Private Sub Menuarch_Click()
   Call Command11_Click
End Sub

Private Sub Orpago_Click()
   Call Orpago1_Click
End Sub

 Sub Orpago1_Click()
   If EXISTE%("ACRGES04.EXE") > 0 Then
     ACONEC$ = "ACRGES04/ORPAGO"
     Call CONECRUN(ACONEC$, "Ordenes de Pago")
   Else
     Call COMUNI("Falta Módulo ACRGES04")
   End If
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cobranzas y Pagos/Ordenes de Pago
  'Call Orpago_Click
End Sub

Private Sub ped_clientes_Click()
  Call CONECRUN("PEDIDO02", "Pedido de Clientes")
End Sub

Private Sub ped_clientes1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/Pedido de Clientes
  Call ped_clientes_Click
End Sub

Private Sub pres_clientes_Click()
   On Error Resume Next
   Call Command9_Click
   On Error GoTo 0
End Sub

Private Sub pres_clientes1_Click()
   On Error Resume Next
   Call Command9_Click
   On Error GoTo 0
End Sub

Private Sub print1_Click()
  GESVAL00.Option2.Value = True ' Opcion diferidos
  GESVAL00.Option4.Value = True ' Opcion por impresora
  Call GESVAL00.Command2_Click ' Ejecuta la consulta
End Sub

Private Sub print12_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Propios Diferidos/Impresora
  Call print1_Click
End Sub

Private Sub print2_Click()
  GESVAL00.Option1.Value = True ' Opcion cartera
  GESVAL00.Option4.Value = True ' Opcion impresora
  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub print22_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Cartera de Valores/Impresora
   Call print2_Click
End Sub

Private Sub Recibo_Click()
   Call Recibo1_Click
End Sub

Private Sub Recibo1_Click()
   If EXISTE%("VENGES04.EXE") > 0 Then
     ACONEC$ = "VENGES04/RECIBO"
     Call CONECRUN(ACONEC$, "Recibo de Cobranza")
   Else
     Call COMUNI("Falta Módulo VENGES04")
   End If
   ' Realiza una llamada a la opcion de menu
   ' Tesoreria/Cobranzas y Pagos/Recibo de Cobranza
   'Call Recibo_Click
End Sub

Private Sub repara_Click()
   Call CONECRUN("REPARA02", "Sub-Sistema de Reparaciones")
End Sub

Private Sub Res_cuenta_Click()
  GESVAL00.Option5.Value = True ' opcion impresora
  Call GESVAL00.Command17_Click ' Ejecuta la consulta
End Sub

Private Sub res_cuenta1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Consultas e Informes\Resumen de Cuenta
  Call Res_cuenta_Click
End Sub

Private Sub Sal_Tes_Click()
  GESVAL00.Option8.Value = True ' Opcion saldos
  GESVAL00.Option5.Value = True ' Opcion pantalla
  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub Sal_Tes1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Saldos de Tesoreria
  Call Sal_Tes_Click
End Sub

Private Sub scr1_Click()
  GESVAL00.Option2.Value = True ' Opcion diferidos
  GESVAL00.Option5.Value = True ' Opcion pantalla
  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub scr11_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Propios Diferidos/Pantalla
  Call scr1_Click
End Sub

Private Sub scr2_Click()
  GESVAL00.Option1.Value = True ' Opcion cartera
  GESVAL00.Option5.Value = True ' Opcion pantalla
  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub scr21_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Cartera de Valores/Pantalla
  Call scr2_Click
End Sub

Private Sub Subdia_CajaBan_Click()
  GESVAL00.Option4.Value = True ' opcion impresora
  Call GESVAL00.Command15_Click ' Ejecuta la consulta
End Sub

Private Sub Subdia_CajaBan1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Consultas e Informes\Subdiario de Caja y Bancos
  Call Subdia_CajaBan_Click
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
   Call MENUTES.Command3_Click
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
        Image1.Picture = LoadPicture(ARCHIMAG$)
        ' Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
        If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") < 1 Then ' PARA CASO GENTILI
           Picture11.Visible = True
           IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
           '
           FORMALTER.Image3.Picture = LoadPicture("")
           FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
           GoTo 12
        End If
        '
      End If
   End If
   '
   Picture11.Visible = False
   If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") > 0 Then
     Picture11.Visible = True
     IMAGEFILE$ = LUDAT$ + "LOGOEMPR.BMP"
     '
     FORMALTER.Image3.Picture = LoadPicture("")
     FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
     '
12   HIMAGEN = FORMALTER.Image3.Picture.Width
     VIMAGEN = FORMALTER.Image3.Picture.Height
     '
     HCUADRO = Picture11.Width
     VCUADRO = Picture11.Height
     '
     REDUH = HIMAGEN / HCUADRO
     REDUV = VIMAGEN / VCUADRO
     REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
     If REDUC < 1 Then REDUC = 1
     '
     With Picture11
       .Height = VCUADRO * REDUV / REDUC
       .Width = HCUADRO * REDUH / REDUC
       .Left = .Left + (HCUADRO - Picture11.Width) / 2
         If .Left < 300 Then Left = 300
       .Top = .Top - (VCUADRO - Picture11.Height) / 2
         If .Top < 400 Then .Top = 400
     End With
     '
     Image1.Width = Picture11.Width
     Image1.Height = Picture11.Height
     Image1.Stretch = True
     Image1.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
     '
     FORMALTER.Image3.Picture = LoadPicture("")
   End If
   '
End Sub

