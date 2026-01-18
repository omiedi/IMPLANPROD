VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INIGES07 
   BackColor       =   &H00C2C2C2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Gestión"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INIGES07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.CommandButton Command19 
      Caption         =   "Sysflom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2640
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2640
      Picture         =   "INIGES07.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Acceso a Seguimiento Comercial"
      Top             =   1320
      Visible         =   0   'False
      Width           =   600
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
      Picture         =   "INIGES07.frx":216C
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
      Picture         =   "INIGES07.frx":2476
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
      Picture         =   "INIGES07.frx":2780
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
      Left            =   3120
      Picture         =   "INIGES07.frx":2A8A
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
      Picture         =   "INIGES07.frx":2ECC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Acreedores - Cuentas por Pagar"
      Top             =   5145
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
      Picture         =   "INIGES07.frx":330E
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
      Picture         =   "INIGES07.frx":3618
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
      Picture         =   "INIGES07.frx":3922
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
      Picture         =   "INIGES07.frx":3C2C
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
         Height          =   700
         Left            =   180
         Picture         =   "INIGES07.frx":B3BE
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cambio de Clave Personal de Acceso"
         Top             =   6370
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
         Picture         =   "INIGES07.frx":B6C8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Sub-Sistema de Presupuestación de Cámaras"
         Top             =   4005
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
         Picture         =   "INIGES07.frx":B9D2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Sub-Sistema de Reparaciones"
         Top             =   3285
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
         Picture         =   "INIGES07.frx":BCDC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Accesos Directos a Otras Aplicaciones"
         Top             =   4710
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
         Picture         =   "INIGES07.frx":BFE6
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
         Picture         =   "INIGES07.frx":C130
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
         Picture         =   "INIGES07.frx":C43A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sistema de Gestion de Stocks"
         Top             =   2580
         Visible         =   0   'False
         Width           =   700
      End
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
         Height          =   700
         Left            =   180
         Picture         =   "INIGES07.frx":C744
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
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
         Picture         =   "INIGES07.frx":CA4E
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   5670
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
         Picture         =   "INIGES07.frx":CE90
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
         Picture         =   "INIGES07.frx":D19A
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
      Picture         =   "INIGES07.frx":D2E4
      ScaleHeight     =   1170
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   720
      Width           =   1995
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
      Height          =   1140
      Left            =   7245
      Picture         =   "INIGES07.frx":14CAA
      ScaleHeight     =   1140
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
      Height          =   1125
      Left            =   1365
      Picture         =   "INIGES07.frx":1C670
      ScaleHeight     =   1125
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   5340
      Width           =   1995
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
      Picture         =   "INIGES07.frx":2416A
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
      Picture         =   "INIGES07.frx":2BDFC
      ScaleHeight     =   1485
      ScaleWidth      =   4800
      TabIndex        =   4
      Top             =   3000
      Width           =   4800
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "INIGES07.frx":4407E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "INIGES07.frx":442B2
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
      Begin VB.Menu PLH151 
         Caption         =   "-"
      End
      Begin VB.Menu ped_clientes 
         Caption         =   "Sub-Sistema de Pedidos de Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu Pre_clientes 
         Caption         =   "Sub-Sistema de Emisión de Presupuestos"
         Enabled         =   0   'False
      End
      Begin VB.Menu PLH15 
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
      Begin VB.Menu contab_gen 
         Caption         =   "Sistema de Contabilidad General"
         Enabled         =   0   'False
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
Attribute VB_Name = "INIGES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Amacli_Click()
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     ACONEC$ = "ARCHIG07/AMACLI"
     Call CONECRUN(ACONEC$, "Maestro de Clientes y Deudores")
   Else
     Call COMUNI("Falta Módulo ARCHIG07")
   End If
'   If EXISTE%("AMACLI02.EXE") > 0 Then
'      Call CONECRUN("AMACLI02", "Maestro de Clientes y Deudores")
'    Else
'      Call CONECRUN("ARCH_GES/AMACLI", "Maestro de Items de Stock")
'   End If
End Sub

Private Sub Amapro_Click()
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     ACONEC$ = "ARCHIG07/AMAPRO"
     Call CONECRUN(ACONEC$, "Maestro de Proveedores y Sub-Contratistas")
   Else
     Call COMUNI("Falta Módulo ARCHIG07")
   End If
'   Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
End Sub

Private Sub Amasto_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If EXISTE%("AMASTO02.EXE") > 0 Then
       Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
     Else
       Call CONECRUN("ARCHIG07/AMASTO", "Maestro de Items de Stock")
   End If
End Sub

Private Sub Bus_che_Click()
  GESVAL07.Option6.Value = True
  Call GESVAL07.Command7_Click
'  GESVAL00.Option6.Value = True
'  Call GESVAL00.Command7_Click
End Sub

Private Sub bus_che1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Busqueda de Cheque
  Call Bus_che_Click
End Sub

Private Sub cgc_acree_Click()
   If EXISTE%("ACRGES07.EXE") > 0 Then
     ACONEC$ = "ACRGES07/CGCOMPRAS"
     Call CONECRUN(ACONEC$, "Control Gestión de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
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
   Call CONECRUN("VENGES07/CGVENTAS", "Control Gestion de Ventas")
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
  If PROCACTI% < 1 Then
    Call GENBACKUP
    Call CANCELUCOM
    Call FINAL("")
  End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call CONECRUN("ORCOMP07", "Sub-Sistema de Compras")
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   'ARCHIM00.Show 1
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("ARCHIG07", "A-B-M-L Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
'   If EXISTE%("VENGES09.EXE") > 0 Then
'      Call CONECRUN("VENGES09", "Gestión Ventas y Deudores")
'      GoTo 99
'   End If
   '
   Call CONECRUN("VENGES07", "Gestión Ventas y Deudores")
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If DERACCE%("GESVALO", "Tesorería y Gestion de Valores") >= 2 Then
      GESVAL07.Show 1
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
  '\\\OT-05-0966-OD-28/10/08///
   Command14.Enabled = False
   If EXISTE%("SEGCOM07.EXE") > 0 Then
       Call CONECRUN("SEGCOM07", "Seguimiento Comercial")
   End If
   Command14.Enabled = True
  '\\\OT-05-0966-OD-fin///
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call CONECRUN("ACRGES07", "Gestión Compras y Acreedores")
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
   Command17.Enabled = False
   Call CONECRUN("FSETUP04/CLAVUSER", "Cambio de Clave Personal de Acceso")
99 Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("CONTAB07.EXE") > 0 Then
       Call CONECRUN("CONTAB07", "Contabilidad General")
     ElseIf EXISTE%("CONTAGEN.EXE") > 0 Then
       Call SELEJER(EJUI%)
       If Abs(EJUI%) <> 1 Then GoTo 99
       Call CONECRUN("CONTAGEN", "Contabilidad General")
   End If
99 Command18.Enabled = True
End Sub

Private Sub Command19_Click()
 Command19.Enabled = False
 FLEXSETUP.Show 1
 Command19.Enabled = True
 
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO07"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   TTXX$ = ""
   Dim ACONE$(16), DECONE$(16)
   ICONE% = 0
   If EXISTE%("FLGEST07.EXE") > 0 Then
      TTXX$ = TTXX$ + "FLEXOFT Industrial\"
      ICONE% = ICONE% + 1: ACONE$(ICONE%) = "FLEXO07": DECONE$(ICONE%) = "Sistema de Gestión Industrial"
   End If
   If EXISTE%("APROYE07.EXE") > 0 Then
     TTXX$ = TTXX$ + "Administacion de Proyectos\"
     ICONE% = ICONE% + 1: ACONE$(ICONE%) = "APROYE07": DECONE$(ICONE%) = "Administración de Proyectos"
   End If
   '
   If EXISTE%("DSGFRM09.EXE") > 0 Then
     TTXX$ = TTXX$ + "Diseñador de Formularios\"
     ICONE% = ICONE% + 1: ACONE$(ICONE%) = "DSGFRM09": DECONE$(ICONE%) = "Diseñador de Formularios"
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
5  Call ACCDIR("FLGEST07")
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
   If DERACCE%("FSETUP07/", "CONFIGURACION GENERAL") < 2 Then GoTo 99
   FSETUPNEW.Show 1
   Call FINAL("")
'   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
'   Call FINAL("")
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If EXISTE%("CONSTO07.EXE") > 0 Then
       Call CONECRUN("CONSTO07", "Control General de Stocks")
     ElseIf EXISTE%("STOBAS07.EXE") > 0 Then
       Call CONECRUN("STOBAS07", "Control Básico de Stocks")
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If DERACCE%("GENREP", "Generador de Reportes") >= 2 Then
    Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 And EXISTE%("COTCLI09.EXE") > 0 Then
     OPX% = ALTERNA%("Emisión de Presupuestos\Pedidos de Clientes\Solicitud de Cotización")
   ElseIf EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 Then
      OPX% = ALTERNA%("Emisión de Presupuestos\Pedidos de Clientes")
   End If
   If OPX% = 1 Then GoTo 10
   If OPX% = 2 Then GoTo 20
   If OPX% = 3 Then GoTo 30
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   '
   If EXISTE%("PRESUP07.EXE") > 0 Then
10    Call CONECRUN("PRESUP07", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
   If EXISTE%("PEDIDO07.EXE") > 0 Then
20    Call CONECRUN("PEDIDO07", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
   If EXISTE%("COTCLI09.EXE") > 0 Then
30    Call CONECRUN("COTCLI09", "Solicitud de Cotización")
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

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("PREFRI07.EXE") > 0 Then
        Call CONECRUN("PREFRI07", "Presupuestación Cámaras Frigoríficas")
      ElseIf EXISTE%("PREFLE07.EXE") > 0 Then
        Call CONECRUN("PREFLE07", "Presupuestación Sistemas FLEXOFT")
      ElseIf EXISTE%("ANACOS07.EXE") > 0 Then
        Call CONECRUN("ANACOS07", "Análisis de Costos Standard")
   End If
   Command9.Enabled = True
End Sub

Private Sub compras_Click()
   Call CONECRUN("ORCOMP07", "Sub-Sistema de Compras")
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
   Call CONECRUN("CONTAB07", "Contabilidad General")
End Sub

Private Sub ctrl_stock_Click()
   Call CONECRUN("CONSTO07", "Control de Stocks")
End Sub

Private Sub Deposito_Click()
   Call GESVAL07.Command11_Click
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Fac_provee_Click()
   If EXISTE%("ACRGES07.EXE") > 0 Then
     ACONEC$ = "ACRGES07/FACTURA"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
   End If
'   Call CONECRUN("ACRE_GES/FACTURA", "Facturacion de Ventas")
End Sub

Private Sub Fac_ven_Click()
   Call CONECRUN("VENGES07/FACTURA", "Facturacion de Ventas")
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
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")
   '
   
   EJER% = 1
   Call GRARAM(69, Chr$(1))
   '
   If EXISTE%("CONTAB07.EXE") > 0 Then
      Command18.Visible = True
      contab_gen.Enabled = True
   End If
   If EXISTE%("CONSTO07.EXE") > 0 Or EXISTE%("STOBAS07.EXE") > 0 Then
      Command5.Visible = True
      ctrl_stock.Enabled = True
   End If
   If EXISTE%("REPARA02.EXE") > 0 Then
     Command8.Visible = True
     repara.Enabled = True
   End If
   If EXISTE%("CONMOD07.EXE") > 0 Then
     Command31.Visible = True
     CONMOD.Enabled = True
   End If
   If EXISTE%("PEDIDO07.EXE") > 0 Then
      Command7.Visible = True
      Command7.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
      ped_clientes.Enabled = True
      ped_clientes1.Enabled = True
   End If
   If EXISTE%("PRESUP07.EXE") > 0 Then
      Command7.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
      Command7.Visible = True
      Pre_clientes.Enabled = True
   End If
      '
   If EXISTE%("COTCLI09.EXE") > 0 Then
      Command7.ToolTipText = "Solicitud de Cotizaciones"
      Command7.Visible = True
   End If
   
   If EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 And EXISTE%("COTCLI09.EXE") > 0 Then
      Command7.ToolTipText = "Sub-Sistemas de Presupuestos, Pedidos de Clientes y Solicitud de Cotizaciones"
   ElseIf EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 Then
      Command7.ToolTipText = "Sub-Sistemas de Presupuestos y Pedidos de Clientes"
  End If
   If EXISTE%("REMITO07.EXE") > 0 Then
     Command23.Visible = True
   End If
   '
   If EXISTE%("ORCOMP07.EXE") > 0 Then
      Command10.Visible = True
      compras.Enabled = True
      compras1.Enabled = True
   End If
   '
   If EXISTE%("ANACOS07.EXE") > 0 Then
     Command9.ToolTipText = "Módulo de Análisis de Costos de Productos"
     Command9.Caption = "An-Cos"
     Command9.Visible = True
   End If
   '
   If EXISTE%("PREFRI07.EXE") > 0 Or EXISTE%("PREFLE07.exe") > 0 Then
     Command9.Visible = True
   End If
   
   '
   If EXISTE%("SEGCOM07.EXE") > 0 Then
      Command14.Visible = True
   End If
   '
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'      Caption = Caption + "  -  " + EPAC$
'   End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
   '
   UTILIZA_SKIN% = 1
'   Call APLICACIONSKINS(Me, Picture7)
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
  MsgBox "FLEXOFT Soluciones Informáticas" & vbCrLf & _
         "Telefax: (00 54 11) 5353-1444 / 5353-1472" & vbCrLf & _
         "Email: asistecflexoft@zetti.com.ar", vbQuestion, _
         "Mesa de Ayuda FLEXOFT"
End Sub

Private Sub Lis_pre_Click()
   If EXISTE%("LPREAC07.EXE") > 0 Then
     Call CONECRUN("LPREAC07", "Lista de Precios")
   Else
     Call COMUNI("Falta Módulo LPREAC07")
   End If
'   Call CONECRUN("VENGES07/LPRECIOS", "Lista de Precios")
End Sub

Private Sub Menuarch_Click()
   Call Command11_Click
End Sub

Private Sub Orpago_Click()
   If EXISTE%("ACRGES07.EXE") > 0 Then
     ACONEC$ = "ACRGES07/FACTURA"
     Call CONECRUN(ACONEC$, "Facturación de Compras")
   Else
     Call COMUNI("Falta Módulo ACRGES07")
   End If
End Sub

Private Sub Orpago1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cobranzas y Pagos/Ordenes de Pago
  Call Orpago_Click
End Sub

Private Sub ped_clientes_Click()
  Call CONECRUN("PEDIDO07", "Pedidos de Clientes")
End Sub

Private Sub pre_clientes_Click()
  Call CONECRUN("PRESUP07", "Presupuestos a Clientes")
End Sub

Private Sub ped_clientes1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Aplicaciones Comerciales/Venta C.Corriente/Pedido de Clientes
  Call ped_clientes_Click
End Sub

Private Sub print1_Click()
  GESVAL07.Option2.Value = True ' Opcion diferidos
  GESVAL07.Option4.Value = True ' Opcion por impresora
  Call GESVAL07.Command2_Click ' Ejecuta la consulta
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
  GESVAL07.Option1.Value = True ' Opcion cartera
  GESVAL07.Option4.Value = True ' Opcion impresora
  Call GESVAL07.Command2_Click ' Ejecuta consulta
'  GESVAL00.Option1.Value = True ' Opcion cartera
'  GESVAL00.Option4.Value = True ' Opcion impresora
'  Call GESVAL00.Command2_Click ' Ejecuta consulta
End Sub

Private Sub print22_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Cartera de Cheques/Cartera de Valores/Impresora
   Call print2_Click
End Sub

Private Sub Recibo_Click()
   If EXISTE%("VENGES07.EXE") > 0 Then
     ACONEC$ = "VENGES07/RECIBO"
     Call CONECRUN(ACONEC$, "Recibo de Cobranza")
   Else
     Call COMUNI("Falta Módulo VENGES07")
   End If
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
'  GESVAL00.Option5.Value = True ' opcion impresora
'  Call GESVAL00.Command17_Click ' Ejecuta la consulta
End Sub

Private Sub res_cuenta1_Click()
  ' Realiza una llamada a la opcion de menu
  ' Tesoreria/Consultas e Informes\Resumen de Cuenta
  Call Res_cuenta_Click
End Sub

Private Sub Sal_Tes_Click()
  GESVAL07.Option8.Value = True ' Opcion saldos
  GESVAL07.Option5.Value = True ' Opcion pantalla
  Call GESVAL07.Command2_Click ' Ejecuta consulta
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
  GESVAL07.Option2.Value = True ' Opcion diferidos
  GESVAL07.Option5.Value = True ' Opcion pantalla
  Call GESVAL07.Command2_Click ' Ejecuta consulta
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
  GESVAL07.Option1.Value = True ' Opcion cartera
  GESVAL07.Option5.Value = True ' Opcion pantalla
  Call GESVAL07.Command2_Click ' Ejecuta consulta
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
   Call GESVAL07.Command10_Click
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(Control$("", "ISOLOGO"))
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
      x$ = TRIM$(UCase$(File1.LIST(U& - 1)))
      If x$ <> "" Then
        If x$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + x$) > 0 Then
            If Left$(x$, 7) = "FLEXO07" Then GoTo 9
            If Left$(x$, 8) = "FLGEST07" Then GoTo 9
            If Left$(x$, 8) = "CONTAB07" Then GoTo 9
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

Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
           '
        Case "GESVAL"
           Call Command13_Click                 'GESTION DE TESORERIA
        Case "GESTRAN"
           Call Command13_Click                 'GESTION DE TESORERIA
        Case "GESCON"
           Call Command13_Click                 'GESTION DE TESORERIA
        Case "BUSCHENU"
           Call GESVAL07.Command7_Click         'BUSCA CHEQUE POR NUMERO
        Case "HISTCHEQ"
            GESVAL07.Option10.Value = True
            Call GESVAL07.Command2_Click        'HISTORICO DE CHEQUES
        Case "DEVALCAR"
            GESVAL07.Option1.Value = True
            Call GESVAL07.Command2_Click        'DETALLE DE CARTERA DE VALORES
        Case "CHENTDEB"
            GESVAL07.Option2.Value = True
            Call GESVAL07.Command2_Click        'CHEQUES ENTREGADOS POR DEBITAR
        Case "REDIATES"
            GESVAL07.Option8.Value = True
            Call GESVAL07.Command2_Click        'RESUMEN DIARIO DE TESORERIA
        Case "DEPOBANC"
            Call GESVAL07.Command11_Click       'DEPOSITO BANCARIO
        Case "TRANFOND"
            Call GESVAL07.Command10_Click       'TRANSFERENCIA DE FONDOS
        Case "RECHSREC"
            Call GESVAL07.Command13_Click       'RECEPCION DE CHEQUES SIN RECIBO DE COBRANZA
        Case "ENCHESORD"
            Call GESVAL07.Command16_Click       'ENTREGA DE VALORES SIN ORDEN DE PAGO
        Case "CONCBANC"
            Call GESVAL07.Command12_Click       'CONCILIACION BANCARIA
        Case "DECUTACR"
            Call GESVAL07.Command20_Click       'DEPOSITO CUPONES TARJETAS DE CREDITO
            Case "CONGEGEST"
            Call GESVAL07.Command9_Click        'CONFIGURACION GESTION DE VALORES
        Case Else
           TERMINA% = 0
   End Select
End Sub

