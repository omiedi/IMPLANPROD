VERSION 5.00
Begin VB.Form INIGEST04 
   BackColor       =   &H00C2C2C2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Gestión"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INIGEST04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
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
      Left            =   3240
      Picture         =   "INIGEST04.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Generación de Presupuestos a Clientes"
      Top             =   5520
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
      Left            =   8760
      Picture         =   "INIGEST04.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   28
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
      Left            =   2760
      Picture         =   "INIGEST04.frx":233E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Sub-Sistema de Pedidos de Clientes"
      Top             =   6240
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
      Picture         =   "INIGEST04.frx":2780
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
      Picture         =   "INIGEST04.frx":2BC2
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
      Picture         =   "INIGEST04.frx":2ECC
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
      Picture         =   "INIGEST04.frx":31D6
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
      Picture         =   "INIGEST04.frx":34E0
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
      Height          =   8600
      Left            =   10800
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   0
      Width           =   1275
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
         Picture         =   "INIGEST04.frx":AC72
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Sub-Sistema de Control de Mano de Obra"
         Top             =   4680
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
         Picture         =   "INIGEST04.frx":AF7C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Sub-Sistema de Reparaciones"
         Top             =   3960
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
         Picture         =   "INIGEST04.frx":B286
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Sistema de Gestion de Stocks"
         Top             =   3240
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Contab"
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
         Picture         =   "INIGEST04.frx":B590
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sistema de Contabilidad General"
         Top             =   2160
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
         Picture         =   "INIGEST04.frx":B89A
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "INIGEST04.frx":BBA4
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "INIGEST04.frx":BEAE
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
         Picture         =   "INIGEST04.frx":C2F0
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "INIGEST04.frx":C5FA
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
      Picture         =   "INIGEST04.frx":C744
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   720
      Width           =   2115
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Trabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   22
         Top             =   315
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7245
      Picture         =   "INIGEST04.frx":1410A
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
      Picture         =   "INIGEST04.frx":1BAD0
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
      Picture         =   "INIGEST04.frx":235CA
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
      Picture         =   "INIGEST04.frx":2B25C
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
Attribute VB_Name = "INIGEST04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Amacli_Click()
    If EXISTE%("AMACLI02.EXE") > 0 Then
       Call CONECRUN("AMACLI02", "Maestro de Clientes y Deudores")
     Else
       Call CONECRUN("ARCH_GES/AMACLI", "Maestro de Items de Stock")
   End If
End Sub

Private Sub Amapro_Click()
   Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
End Sub

Private Sub Amasto_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If EXISTE%("AMASTO02.EXE") > 0 Then
       Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
     Else
       Call CONECRUN("ARCH_GES/AMASTO", "Maestro de Items de Stock")
   End If
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
   Call CONECRUN("ACRE_GES/CGCOMPRAS", "Control Gestion de Compras")
End Sub

Private Sub cgc_com1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/
  ' Control Gestion de Compras
  Call cgc_acree_Click
End Sub

Private Sub cgc_ven_Click()
   Call CONECRUN("VENT_GES/CGVENTAS", "Control Gestion de Ventas")
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
     Call CONECRUN("ORTRAB04", "Ordenes de Trabajo")
   Else
     Call COMUNI("Falta módulo ARCHIG04.EXE")
   End If
   'ARCHIM00.Show 1
   'Call CIERRARCH("*.*")
   'Call BAJALDISCO
   'Call CONECRUN("ARCH_GES", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("VENT_GES", "Gestión Ventas y Deudores")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("ORTRAB04.EXE") > 0 Then
     Call CONECRUN("ORTRAB04", "Ordenes de Trabajo")
   Else
     Call COMUNI("Falta módulo ORTRAB04.EXE")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("ACRE_GES", "Gestión Compras y Acreedores")
   Command15.Enabled = True
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

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CONECRUN("CONSTO02", "Control Básico de Stocks")
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
   Call CONECRUN("ORCOMP03", "Sub-Sistema de Compras")
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
   Call CONECRUN("ACRE_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Fac_ven_Click()
   Call CONECRUN("VENT_GES/FACTURA", "Facturacion de Ventas")
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
   Call CONECRUN("VENT_GES/LPRECIOS", "Lista de Precios")
End Sub

Private Sub Menuarch_Click()
   Call Command11_Click
End Sub

Private Sub Orpago_Click()
   Call MENUTES.Command22_Click
End Sub

Private Sub Orpago1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cobranzas y Pagos/Ordenes de Pago
  Call Orpago_Click
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
   Call MENUTES.Command7_Click
End Sub

Private Sub Recibo1_Click()
   ' Realiza una llamada a la opcion de menu
   ' Tesoreria/Cobranzas y Pagos/Recibo de Cobranza
   Call Recibo_Click
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

