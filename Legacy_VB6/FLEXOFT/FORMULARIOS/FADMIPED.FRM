VERSION 5.00
Begin VB.Form FADMIPED 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Pedidos"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FADMIPED.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
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
      Left            =   5565
      Picture         =   "FADMIPED.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   315
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   5145
      ScaleHeight     =   795
      ScaleWidth      =   90
      TabIndex        =   59
      Top             =   340
      Width           =   150
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asignación de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5145
      TabIndex        =   39
      Top             =   3780
      Width           =   4635
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   61
         Top             =   525
         Width           =   1800
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
            Picture         =   "FADMIPED.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Analisis de Pedidos de Kits"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis de Pedidos   de Kits"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   700
            TabIndex        =   63
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   47
         Top             =   525
         Width           =   1800
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
            Left            =   0
            Picture         =   "FADMIPED.frx":0460
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Asignación de Materiales y Ordenes de Despacho "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Ordenes de Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   700
            TabIndex        =   49
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   40
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command36 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   46
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   45
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   44
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   43
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   42
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   41
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Archivos Maestros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   210
      TabIndex        =   14
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   25
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command35 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   31
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   30
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   29
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   28
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   27
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   26
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Clientes / Proveedores"
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
         Left            =   315
         TabIndex        =   20
         Top             =   1890
         Width           =   4005
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   21
            Top             =   420
            Width           =   3165
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
               Left            =   -20
               Picture         =   "FADMIPED.frx":076A
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
               Top             =   0
               Width           =   560
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
               Left            =   2520
               Picture         =   "FADMIPED.frx":0BAC
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Base de Datos de Proveedores y Subcontratistas"
               Top             =   0
               Width           =   600
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes, Proveedo- res y Subcontratistas "
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
               TabIndex        =   24
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Artículos y Rutas"
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
         Left            =   315
         TabIndex        =   15
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   16
            Top             =   420
            Width           =   3165
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
               Picture         =   "FADMIPED.frx":0EB6
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
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
               Left            =   1950
               Picture         =   "FADMIPED.frx":11C0
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Fórmulas y Estructuras de Conjunto"
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton Command32 
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
               Picture         =   "FADMIPED.frx":14CA
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Hojas de Ruta - Secuencia de Operaciones de Fabricación"
               Top             =   0
               Width           =   600
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestros de Items de Stock"
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
               TabIndex        =   19
               Top             =   105
               Width           =   1485
            End
         End
      End
   End
   Begin VB.CommandButton Command14 
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
      Height          =   855
      Left            =   6432
      Picture         =   "FADMIPED.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command8 
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
      Height          =   855
      Left            =   7300
      Picture         =   "FADMIPED.frx":1C16
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command29 
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
      Left            =   8165
      Picture         =   "FADMIPED.frx":1F20
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
      Width           =   750
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
      Height          =   855
      Left            =   9030
      Picture         =   "FADMIPED.frx":222A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base de Datos de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   2
      Top             =   3780
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   900
         ScaleHeight     =   585
         ScaleWidth      =   2775
         TabIndex        =   33
         Top             =   525
         Width           =   2830
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
            Left            =   550
            Picture         =   "FADMIPED.frx":2374
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Repetir Impresion de Pedido Registrado"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
         Begin VB.CommandButton Command42 
            Enabled         =   0   'False
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
            Left            =   550
            Picture         =   "FADMIPED.frx":29DE
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Registrar Pedido Interno"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command43 
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
            Left            =   2220
            Picture         =   "FADMIPED.frx":2E20
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Revertir Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command41 
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
            Picture         =   "FADMIPED.frx":2F6A
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Registrar Pedido de Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command46 
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
            Left            =   1110
            Picture         =   "FADMIPED.frx":33AC
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Cambio de Cantidades y Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command45 
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
            Left            =   1680
            Picture         =   "FADMIPED.frx":36B6
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   3
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   5145
      TabIndex        =   1
      Top             =   1470
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   420
         TabIndex        =   56
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2205
            TabIndex        =   58
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   525
            TabIndex        =   57
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   50
         Top             =   525
         Width           =   3795
         Begin VB.CommandButton Command63 
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
            Left            =   1950
            Picture         =   "FADMIPED.frx":39C0
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Detalle por Número de Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
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
            Picture         =   "FADMIPED.frx":3CCA
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Encabezados de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command62 
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
            Left            =   2550
            Picture         =   "FADMIPED.frx":3FD4
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Pedidos Pendientes por Cliente"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command64 
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
            Picture         =   "FADMIPED.frx":4416
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Pedidos Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Datos de Pedidos"
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
            TabIndex        =   54
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2730
      TabIndex        =   0
      Top             =   5460
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "FADMIPED.frx":4720
      Stretch         =   -1  'True
      Top             =   5355
      Width           =   10920
   End
End
Attribute VB_Name = "FADMIPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command14_Click()
    OPORDESPA.Show 1
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call CONECRUN("AMACLI", "Padrón Maestro de Clientes")
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Padrón de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    Command2.Enabled = True
End Sub

Private Sub Command32_Click()
    Command32.Enabled = False
    Call CONECRUN("HORUTAS", "Secuencia de Operaciones Standard")
    Command32.Enabled = True
End Sub

Private Sub Command4_Click()
    If InStr(UCase$(EMPREAC$), "SABO") < 1 Then
         Call OPNOIN("Análisis de Pedidos de Kits")
         Exit Sub
       Else
         FANAKITS.Show 1
    End If
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Revertir Anulación de Pedido Itemizado")
       Exit Sub
    End If
    Call BLOQARCH("PEDICLIE")
    Call DESANUPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Command43.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call FINAL("*CPROGEN")
    Command25.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("FORMULAS", "Fórmulas y Estructuras de Producto")
    Command3.Enabled = True
End Sub

Private Sub Command42_Click()
    Command42.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
         Call MENSERR(24, "Imposible Registrar Pedido Interno\para Formato Itemizado. Consulte.")
       Else
         Call PEDINTNUE
    End If
    '
    Command42.Enabled = True
End Sub

Private Sub Command41_Click()
    Command41.Enabled = False
    '
    EPAC$ = TRIM$(EMPREAC$)
    If InStr(EPAC$, "NOVOESPU") > 0 Then
       Call CARPEDNOVO
       GoTo 99
    End If
    '
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
         FORCARPED.Show 1
       Else
         Call BLOQARCH("PEDICLIE")
         Call PEDCLINUE
         Call CIERRARCH("*.*")
         Call LIBARCH("PEDICLIE")
    End If
    '
99  Command41.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   Call COPEDPE
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call REPIMPED             ' repetir impresion de pedido
   Command6.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
        Call LIPEDIT(4)
        GoTo 99
    End If
    '
    If Option1.Value = True Then
           SALI$ = "*COPEDMA"
           If FORMAPED% = 2 Then SALI$ = "*COQEDMA"
           Call FINAL(SALI$)
        Else
           Screen.MousePointer = 11
           JJ& = 0
           FIN& = ULTREG&("PEDDETP")
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
           If MODO% < 1 Then GoTo 99
           FLI1% = -1: FLI2% = FINMESACT%
           If MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           VECPEDI$ = ""
           For U& = 1 To FIN&
               Call TRACE(20, U&, FIN&)
               X$ = REGLEIDO$("PEDDETP", U&)
               FENTRESOL% = CVI(Mid$(X$, 25, 2))
               If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                     Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                     Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                     Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                     '
                     CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                     If CAPEN < 0 Then CAPEN = 0
                     CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                     If CANOASI < 0 Then CANOASI = 0
                     '
                     NC% = CVI(Mid$(X$, 7, 2))
                     PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                     If OPX% <> 2 Then PREUNI = 0
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     INOAS# = CDbl(CANOASI * PREUNI)
                     
                     Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                     Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                     Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                     Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                     Call REPLA(Y$, Space$(34), 127, 34)
                     '
                     JJ& = JJ& + 1
                     Call GRAREG("LIPEDPE", Y$, JJ&)
                     VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                   End If
                End If
           Next U&
           '
           FIN& = ULTREG&("ORDESPA")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ORDESPA", U&)
             CADESPAX = CVS(Mid$(X$, 79, 4))
             If CADESPAX >= 0.05 Then
               ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
               For W& = 1 To Len(VECPEDI$) / 8
                 If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                   Y$ = REGLEIDO$("LIPEDPE", W&)
                   ODX$ = TRIM$(Mid$(Y$, 127, 34))
                   If ODX$ = "" Then
                       ODX$ = "OD"
                     Else
                       ODX$ = ODX$ + ","
                   End If
                   ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
                   Call REPLA(Y$, ODX$, 127, 34)
                   '
                   If MODO% = 4 Then
                      CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
                      CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
                      Call REPLA(Y$, MKS$(CAENTX), 97, 4)
                      Call REPLA(Y$, MKS$(CASALX), 101, 4)
                   End If
                   '
                   Call GRAREG("LIPEDPE", Y$, W&)
                   GoTo 19
                 End If
               Next W&
             End If
19         Next U&
           '
           FIN& = ULTREG&("ODESEMI")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ODESEMI", U&)
             ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
             For W& = 1 To Len(VECPEDI$) / 8
               If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                 Y$ = REGLEIDO$("LIPEDPE", W&)
                 ODX$ = TRIM$(Mid$(Y$, 127, 34))
                 If ODX$ = "" Then
                    CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
                      If Val(CASIGNA$) > 0.05 Then
                        Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                        Call GRAREG("LIPEDPE", Y$, W&)
                      End If
                 End If
                 GoTo 29
               End If
             Next W&
29         Next U&
           '
           If MODO% = 4 Then
             KK& = JJ&: JJ& = 0
             For U& = 1 To KK&
               Y$ = REGLEIDO$("LIPEDPE", U&)
               If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
                 JJ& = JJ& + 1
                 Call GRAREG("LIPEDPE", Y$, JJ&)
               End If
             Next U&
             Call HEADERS("LIPEDPE", "")
             Call HEADERS("LIPEDPE", "Corresponde a: Pedidos Pendientes sin Orden de Despacho")
           End If
           '
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENMA"
           If FORMAPED% = 2 Then SALI$ = "*SOQENMA"
           Call FINAL(SALI$)
    End If
99  Screen.MousePointer = 1
    Command64.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ADMIPED")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call BLOQARCH("PEDICLIE")
    Call BLOQARCH("FACTURA")
    ORDESPACH.Show 1
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Call LIBARCH("FACTURA")
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Command45.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Anulación de Pedido Itemizado")
       GoTo 99
    End If
    Call BLOQARCH("PEDICLIE")
    Call BLOQARCH("ODESEMI")
    Call ANULAPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
    Call LIBARCH("ODESEMI")
99  Command45.Enabled = True
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call MENSERR(24, "Imposible Modificar Pedido\para Formato Itemizado.\ \Anule y Cargue un Nuevo Pedido.")
       GoTo 99
    End If
    Call BLOQARCH("PEDICLIE")
    Call CAMBIAPED
    Call CIERRARCH("*.*")
    Call LIBARCH("PEDICLIE")
99  Command46.Enabled = True
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    If Option1.Value = True Then
           Call OPNOIN("")     '  Call FINAL("*COPEDNU")
        Else
           Call OPNOIN("")
    End If
    Command61.Enabled = True
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
    If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
        Call LIPEDIT(3)
        GoTo 99
    End If
    '
    If Option1.Value = True Then
           SALI$ = "*COPEDCL"
           If FORMAPED% = 2 Then SALI$ = "*COQEDCL"
           Call FINAL(SALI$)
        Else
           Screen.MousePointer = 11
           JJ& = 0
           FIN& = ULTREG&("PEDDETP")
           '
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores\Total\Sin O/Despacho")
           If MODO% < 1 Then GoTo 99
           FLI1% = -1: FLI2% = FINMESACT%
           If MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
             ElseIf MODO% = 3 Or MODO% = 4 Then
               FLI1% = -1: FLI2% = 32767
           End If
           '
           OPX% = ALTERNA%("Solo Cantidades\Cantidades y Precios")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           '
           VECPEDI$ = ""
           For U& = 1 To FIN&
                Call TRACE(20, U&, FIN&)
                X$ = REGLEIDO$("PEDDETP", U&)
                FENTRESOL% = CVI(Mid$(X$, 25, 2))
                If FENTRESOL% >= FLI1% Then
                   If FENTRESOL% <= FLI2% Then
                      Y$ = String$(46, 0) + Left$(X$, 6) + Mid$(X$, 9, 24) + Mid$(X$, 75)
                      Call REPLA(Y$, Mid$(X$, 7, 2), 1, 2)
                      Call REPLA(Y$, Mid$(X$, 7, 2), 45, 2)
                      Call REPLA(Y$, Mid$(X$, 33, 42), 3, 42)
                      '
                      CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                      If CAPEN < 0 Then CAPEN = 0
                      CANOASI = CAPEN - CVS(Mid$(X$, 75, 4))
                      If CANOASI < 0 Then CANOASI = 0
                      '
                      NC% = CVI(Mid$(X$, 7, 2))
                      PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                      If OPX% <> 2 Then PREUNI = 0
                      ITOTA# = CDbl(CAPEN * PREUNI)
                      INOAS# = CDbl(CANOASI * PREUNI)
                      '
                      Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                      Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                      Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                      Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                      Call REPLA(Y$, Space$(34), 127, 34)
                      '
                      JJ& = JJ& + 1
                      Call GRAREG("LIPEDPE", Y$, JJ&)
                      VECPEDI$ = VECPEDI$ + Mid$(Y$, 47, 4) + Mid$(Y$, 3, 2) + Mid$(Y$, 69, 2)
                   End If
                End If
           Next U&
           '
           FIN& = ULTREG&("ORDESPA")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ORDESPA", U&)
             CADESPAX = CVS(Mid$(X$, 79, 4))
             If CADESPAX >= 0.05 Then
               ABUS$ = Mid$(X$, 7, 4) + Mid$(X$, 37, 2) + Mid$(X$, 31, 2)
               For W& = 1 To Len(VECPEDI$) / 8
                 If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                   Y$ = REGLEIDO$("LIPEDPE", W&)
                   ODX$ = TRIM$(Mid$(Y$, 127, 34))
                   If ODX$ = "" Then
                       ODX$ = "OD"
                     Else
                       ODX$ = ODX$ + ","
                   End If
                   ODX$ = ODX$ + " " + TRIM$(Str$(CVS(Left$(X$, 4))))
                   Call REPLA(Y$, ODX$, 127, 34)
                   '
                   If MODO% = 4 Then
                      CADESPAX = CVS(Mid$(X$, 79, 4))
                      CAENTX = CVS(Mid$(Y$, 97, 4)) + CADESPAX
                      CASALX = CVS(Mid$(Y$, 101, 4)) - CADESPAX
                      Call REPLA(Y$, MKS$(CAENTX), 97, 4)
                      Call REPLA(Y$, MKS$(CASALX), 101, 4)
                   End If
                   '
                   Call GRAREG("LIPEDPE", Y$, W&)
                   GoTo 19
                 End If
               Next W&
             End If
19         Next U&
           '
           FIN& = ULTREG&("ODESEMI")
           For U& = 1 To FIN&
             X$ = REGLEIDO$("ODESEMI", U&)
             ABUS$ = Mid$(X$, 1, 4) + Mid$(X$, 11, 2) + Mid$(X$, 9, 2)
             For W& = 1 To Len(VECPEDI$) / 8
               If Mid$(VECPEDI$, 8 * W& - 7, 8) = ABUS$ Then
                 Y$ = REGLEIDO$("LIPEDPE", W&)
                 ODX$ = TRIM$(Mid$(Y$, 127, 34))
                 If ODX$ = "" Then
                     CASIGNA$ = TRIM$(CSTRING$(Mid$(X$, 17, 4), 3, 10, 1, 2))
                     If Val(CASIGNA$) >= 0.05 Then
                        Call REPLA(Y$, "Asignar " + CASIGNA$, 127, 34)
                        Call GRAREG("LIPEDPE", Y$, W&)
                     End If
                 End If
                 GoTo 29
               End If
             Next W&
29         Next U&
           '
           If MODO% = 4 Then
             KK& = JJ&: JJ& = 0
             For U& = 1 To KK&
               Y$ = REGLEIDO$("LIPEDPE", U&)
               If CVS(Mid$(Y$, 101, 4)) >= 0.05 Then
                 JJ& = JJ& + 1
                 Call GRAREG("LIPEDPE", Y$, JJ&)
               End If
             Next U&
             Call HEADERS("LIPEDPE", "")
             Call HEADERS("LIPEDPE", "Corresponde a: Pedidos Pendientes sin Orden de Despacho")
           End If
           '
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENCL"
           If FORMAPED% = 2 Then SALI$ = "*SOQENCL"
           Call FINAL(SALI$)
    End If
99  Screen.MousePointer = 1
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
        Call LIPEDIT(2)
        GoTo 99
    End If
    '
    If Option1.Value = True Then
           SALI$ = "*COPEDNO"
           If FORMAPED% = 2 Then SALI$ = "*COQEDNO"
           Call FINAL(SALI$)
        Else
           Call OPNOIN("")
    End If
99  Command63.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EMPREAC$, "MAINIERI") > 0 Or InStr(EMPREAC$, "SEALU") > 0 Then
       FAPEDMAI.Show
       Unload FADMIPED
       Exit Sub
       '
       Command42.Visible = False
       Command6.Visible = True
    End If
    '
    'Command42.Visible = False
    'Command6.Visible = True
    '
    If InStr(EMPREAC$, "SABO") < 1 Then
       Command4.Enabled = False
       If InStr(EMPREAC$, "NEUMANN") < 1 Then
          Command9.Enabled = False
       End If
    End If
    '
    'If EXISTE%("PEDNOVO.EXE") > 0 Then
    '    MODO% = ALTERNA%("Procesar Pedidos de Clientes\Control de la Base de Datos")
    '    If MODO% = 1 Then
    '       Call CONECRUN("#PEDNOVO", "Pedidos de Clientes - Novoespuma")
    '      ElseIf MODO% = 2 Then
    '        Call COPEDPE
    '    End If
    '    Call FINAL("")
    'End If
    '
    COMMA$ = Command$
    PPX% = InStr(COMMA$, "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    '
    If TRIM$(COMMA$) = "REVIPEPE" Then
        REVIPEPE% = 1
        Show
        Refresh
        ORDESPACH.Show 1
    End If
    '
End Sub

Sub CARPEDNOVO()
      '
1000  NRO& = 1
      If ULTREG&("PEDENCA") > 0 Then
         X$ = REGLEIDO$("PEDENCA", ULTREG&("PEDENCA"))
         If Len(X$) >= 4 Then NRO& = 1 + CVS(Left$(X$, 4))
      End If
      '
1100  TIPED% = ALTERNA%("Pedido de Cliente\Pedido Interno")
      If TIPED% < 1 Or TIPED% > 2 Then
         Call CIERRARCH("*.*")
         Exit Sub
      End If
      '
      If TIPED% = 2 Then
         PEDX$ = MKS$(NRO&) + MKI$(HOY(HO$)) + MKI$(9999)
         Y$ = Space$(344)
         Call REPLA(Y$, Left$(PEDX$, 8), 1, 8)
         Call REPLA(Y$, "INTERNO", 9, 32)
         '
         Y$ = Y$ + MKI$(HOY(HO$)) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + String$(20, 0)
         Call REPLA(Y$, MKI$(0), 345, 2)
         PEDY$ = Y$
         GoTo 1190
      End If
      '
      If PEDX$ = "" Then PEDX$ = MKS$(NRO&) + MKI$(HOY(HO$)) + MKI$(NC%)
      PEDX$ = OBJPLA$("PEDIDO-NUE", PEDX$)
      If Len(PEDX$) < 24 Then
         Call CIERRARCH("*.*")
         Exit Sub
      End If
      '
      NC% = CVI(Mid$(PEDX$, 7, 2))
      If NC% < 1 Then GoTo 1000
      '
      Call NOTACLI(NC%, 1)
      '
      Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(344)
      Call REPLA(Y$, Left$(PEDX$, 8), 1, 8)
      Call REPLA(Y$, Mid$(PEDX$, 9, 25), 137, 16)
      Call REPLA(Y$, RASOCLI$(NC%), 9, 32)
      Call REPLA(Y$, DOMICLI$(NC%), 41, 32)
      CLOCA$ = TRIM$(CPOSCLI$(NC%))
      If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
      Call REPLA(Y$, CLOCA$ + LOCACLI$(NC%), 73, 32)
      Call REPLA(Y$, TRIM$(CUITCLI$(NC%)) + " " + TRIM$(POSIVAC$(NC%)), 105, 32)
      '
      Y$ = Y$ + MKI$(HOY(HO$)) + Chr$(0) + Chr$(LIPRCLI%(NC%)) + Chr$(CPAGCLI%(NC%)) + Chr$(VENDCLI%(NC%)) + String$(20, 0)
      PEDY$ = OBJPLA$("PEDIDO-ENCA", Y$)
      If PEDY$ = "" Then GoTo 1100
      '
      Call REPLA(PEDY$, MKI$(0), 345, 2)
1110  PEDY$ = OBJPLA$("PEDIDO-ENCB", PEDY$)
      If PEDY$ = "" Then GoTo 1100
      '
      LPRE% = Asc(Mid$(PEDY$, 348, 1))
      If LPRE% < 1 Then GoTo 1110
      '
      ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      '
1190  TIPE% = ALTERNA%("Espumas\Productos")
      If TIPE% < 1 Or TIPE% > 2 Then
         Call CIERRARCH("*.*")
         Exit Sub
      End If
      '
1200  FEPED% = HOY(HOS$)
      ACAR$ = "CARDETS"
      If TIPE% = 2 Then ACAR$ = "CARDETP"
      If TIPED% = 2 Then ARCHIPRE$ = "MASTER"
      XXX% = VENTABU%(ACAR$ + "/NC/ARCHECO(1)=" + ARCHIPRE$)
      '
      If XXX% >= 0 Then GoTo 1300
      PEDX$ = ""
      Call BLANARCH("!CARDETPE")
      Call CIERRARCH("*.*")
      GoTo 1000
      '
1300  If ULTREG&("!CARDETPE") < 1 Then
          GoTo 1200
      End If
      '
      '*************************** Registrando Pedido ...
      '
      ' aqui anular pedidos pendientes
      '
      ' recalcular precios en cardetpe
      '
      ATX% = 0: ATY% = 0: CIANT% = 0
      For U& = 1 To ULTREG&("!CARDETPE")
         X$ = REGLEIDO$("!CARDETPE", U&)
         CI% = CVI(Left$(X$, 2))
         CAN = CVS(Mid$(X$, 47, 4))
         If CAN > 0.05 Then
            ATX% = 1
            If CI% = CIANT% Then GoTo 1310
            CIANT% = CI%
            '
            If TIPE% = 1 Then
               MECU = CVS(Mid$(X$, 27, 4)) * CVS(Mid$(X$, 31, 4)) * CVS(Mid$(X$, 35, 4))
               If MECU < 1 Then ATY% = 1
            End If
            '
            PRELIS = 0: PREUNI = 0
            If ECOARCH$(ARCHIPRE$, MKI$(CI%)) <> "" Then
               Y$ = REGACT$(ARCHIPRE$)
               PRELIS = CVD(Mid$(Y$, 62, 8))
            End If
            '
            PREUNI = PRELIS
            If TIPE% = 1 Then
               PREUNI = PRELIS / 10 ^ 9 * MECU
            End If
1310        TOT# = CDbl(CAN * PREUNI)
            Call REPLA(X$, MKS$(PRELIS), 39, 4)
            Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
            Call GRAREG("!CARDETPE", X$, U&)
            '
         End If
         '
      Next U&
      '
      If ATX% < 1 Or ATY% > 0 Then
          GoTo 1200
      End If
      '
      REFE$ = " - " + TRIM$(Mid$(PEDY$, 137, 16))
      If REFE$ = " - " Then REFE$ = ""
      FEX = CVI(Mid$(PEDY$, 5, 2))
      REFE$ = REFE$ + " - " + FECHATEX$(FEX)
      LOMA% = 48 - Len(REFE$)
      If TIPED% = 2 Then
           REFE$ = "INTERNO " + REFE$
         Else
           REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NC%))), LOMA%) + REFE$
      End If
      '
      PDC$ = REGBLAN$("PEDENCA")
      Call REPLA(PDC$, Left$(PEDY$, 4), 1, 4)
      Call REPLA(PDC$, REFE$, 5, 48)
      Call REPLA(PDC$, Mid$(PEDY$, 5, 4), 53, 4)
      Call REPLA(PDC$, Mid$(PEDY$, 137, 214), 57, 214)
      Call REPLA(PDC$, MKI$(0), 265, 2)
      Call REPLA(PDC$, Chr$(1), 267, 1)
      Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
      Call REPLA(PDC$, String$(14, 0), 275, 14)
      Call REPLA(PDC$, CAVERI$, 289, 14)
      '
      PEDENCA$ = PDC$
      Call REGAPP("PEDENCA", PDC$)
      REGENCA& = ULTREG&("PEDENCA")
      '
      PRIMO& = 1 + ULTREG&("PEDDETA")
      For U& = 1 To ULTREG&("!CARDETPE")
         X$ = REGLEIDO$("!CARDETPE", U&)
         CAN = CVS(Mid$(X$, 47, 4))
         If CAN > 0.05 Then
            PDC$ = REGBLAN$("PEDDETA")
            Call REPLA(PDC$, Left$(PEDY$, 8), 1, 8)
            Call REPLA(PDC$, Mid$(PEDY$, 137, 16), 9, 16)
            Call REPLA(PDC$, MKI$(FEPED%), 25, 2)
            Call REPLA(PDC$, Chr$(1), 27, 1)
            Call REPLA(PDC$, Mid$(PEDY$, 348, 3), 28, 3)
            Call REPLA(PDC$, MKI$(TIPE%), 31, 2)
            Call REPLA(PDC$, Left$(X$, 62), 33, 62)
            Call REPLA(PDC$, String$(18, 0), 95, 18)
            Call REPLA(PDC$, MKS$(CAN), 99, 4)
            Call REPLA(PDC$, MKS$(REGENCA&), 109, 4) ' puntero a PEDENCA
            Call REPLA(PDC$, CAVERI$, 113, 16)
            '
            Call REGAPP("PEDDETA", PDC$)
            '
            RGX& = ULTREG&("PEDDETA")
            Call REPLA(PDC$, MKS$(RGX&), 109, 4)    ' puntero a PEDDETA
            Call REGAPP("PEDDETP", PDC$)
            '
         End If
         '
      Next U&
      '
      GoSub PRIPEDNOVO
      '
      PEDX$ = ""
      Call BLANARCH("!CARDETPE")
      Call CIERRARCH("*.*")
      Call LIBARCH("PEDENCA")
      Call LIBARCH("PEDDETA")
      Call LIBARCH("PEDDETP")
      '
      ' carga de maquinas
      ' *****************
      '
      Screen.MousePointer = 11
      '
1405  FIN& = ULTREG&("PEDDETA")
      QQ% = 0: NROANT& = 0
      For I& = PRIMO& To FIN&
         X$ = REGLEIDO$("PEDDETA", I&)
         NROPED& = CVS(Left$(X$, 4))
         CAPENDE = CVS(Mid$(X$, 99, 4))
         If I& = PRIMO& Then
            FEPED% = CVI(Mid$(X$, 5, 2))
            PRIORI% = 0: STATI% = 1
            NCLI% = CVI(Mid$(X$, 7, 2))
            DESTA$ = Left$(TEL2CLI$(NCLI%), 8)
            RASO$ = RASOCLI$(NCLI%)
            If TRIM$(DESTA$) = "" Then DESTA$ = Left$(TRIM$(RASO$), 8)
            DESTI$ = AJUSTI$(DESTA$, 8)
            For KKK& = 1 To ULTREG&("DESTFAB")
               If Left$(REGLEIDO$("DESTFAB", KKK&), 8) = DESTI$ Then
                  GoTo 1410
               End If
            Next KKK&
            DESTO$ = Space$(80)
            Call REPLA(DESTO$, DESTI$, 1, 8)
            Call REPLA(DESTO$, RASO$, 9, 32)
            Call REGAPP("DESTFAB", DESTO$)
            '
1410     End If
         '
         QQ% = I& - PRIMO& + 1
         NOFAN$ = TRIM$(Str$(NROPED&))
         While Len(NOFAN$) < 4: NOFAN$ = "0" + NOFAN$: Wend
         NOFA$ = "OF-" + NOFAN$ + "-" + TRIM$(Str$(QQ%))
         '
         CI% = CVI(Mid$(X$, 33, 2))
         COD$ = TRIM$(CODEXT$(CI%))
         DE$ = TRIM$(Left$(DESCRIT$(CI%), 24))
         LG$ = "": If CVS(Mid$(X$, 59, 4)) > 1 Then LG$ = TRIM$(Str$(CVS(Mid$(X$, 59, 4))))
         AN$ = "": If CVS(Mid$(X$, 63, 4)) > 1 Then AN$ = TRIM$(Str$(CVS(Mid$(X$, 63, 4))))
         EP$ = "": If CVS(Mid$(X$, 67, 4)) > 1 Then EP$ = TRIM$(Str$(CVS(Mid$(X$, 67, 4))))
         MEDI$ = "": If LG$ + AN$ + EP$ <> "" Then MEDI$ = LG$ + " x " + AN$ + " x " + EP$
         MEDI$ = TRIM$(MEDI$)
         YAECO% = 0
         '
         UNI$ = UNIMED$(CI%)
         REFE$ = COD$ + " - "
         If MEDI$ <> "" Then
              REFE$ = REFE$ + MEDI$
            Else
              REFE$ = REFE$ + TRIM$(DESCRIT$(CI%))
         End If
         CANS$ = " (" + TRIM$(Str$(CAPENDE)) + ")"
         REFE$ = Left$(REFE$, 52 - Len(CANS$)) + CANS$
         '
         Z$ = REGBLAN$("ORDEFABR")
         Call REPLA(Z$, MKS$(NROPED&), 1, 4)
         Call REPLA(Z$, REFE$, 5, 64)
         Call REPLA(Z$, NOFA$, 69, 12)
         Call REPLA(Z$, MKI$(FEPED%), 81, 2)
         Call REPLA(Z$, MKI$(CI%), 83, 2)
         Call REPLA(Z$, UNI$, 85, 4)
         Call REPLA(Z$, MKD$(CDbl(CAPENDE)), 89, 8)
         Call REPLA(Z$, DESTI$, 97, 8)
         Call REPLA(Z$, MKI$(FEPED%), 105, 2)    ' fecha solicitada de entrega
         Call REPLA(Z$, Chr$(PRIORI%), 107, 1)      ' sin prioridad
         Call REPLA(Z$, Chr$(STATI%), 108, 1) ' status = lanzada
         Call REPLA(Z$, MKI$(FEPED%), 109, 2)    ' fecha de lanzamiento
         Call REPLA(Z$, MKS$(0), 113, 4)   ' sector productivo
         Call REPLA(Z$, MKD$(0), 119, 8)
         Call REPLA(Z$, String$(30, 0), 159, 30)
         '
         ZZ$ = AJUSTI$(NOFA$, 12) + REFE$
         '
         Call REGAPP("ORDEFABR", Z$)
         J& = ULTREG&("ORDEFABR")
         Call GRAREG("ECORDEF", ZZ$, J&)
         Call REGAPP("ECORDEP", ZZ$)
         '
         Call GENRESER(NOFA$)
         '
         HAYSEC% = 0
         For TTL& = 1 To ULTREG&("SECOPER")
         SCP$ = REGLEIDO$("SECOPER", TTL&)
         If CVI(Left$(SCP$, 2)) = CI% Then
              NO% = CVI(Mid$(SCP$, 3, 2))
              DOPE$ = Mid$(SCP$, 4, 48)
              EQ1$ = Mid$(SCP$, 53, 6)
              EQAUX$ = Mid$(SCP$, 103, 12)
              LOTESTA = CVS(Mid$(SCP$, 121, 4)): If LOTESTA < 1 Then LOTESTA = 1
              CALOTES = 1 + Int((CAPENDE - 1) / LOTESTA)
              HORAPREP = CVS(Mid$(SCP$, 59, 4)) * CALOTES
              MINLOTE = CALOTES * 60 * CVS(Mid$(SCP$, 63, 4))
              MINLOTE = MINLOTE + (1 + CVS(Mid$(SCP$, 75, 4)) / 100) * CAPENDE * CVS(Mid$(SCP$, 67, 4))
              MINLOTE = MINLOTE + CVS(Mid$(SCP$, 79, 4))
              HORAPROD = MINLOTE / 60
              CAOPS = CVS(Mid$(SCP$, 83, 4))
              '
              Y$ = REGBLAN$("OPERFAB")
              Call REPLA(Y$, NOFA$, 1, 12)
              Call REPLA(Y$, MKI$(CI%), 13, 2)
              Call REPLA(Y$, MKI$(NO%), 15, 2)
              Call REPLA(Y$, DOPE$, 17, 48)
              Call REPLA(Y$, DESTI$, 65, 8)
              Call REPLA(Y$, EQ1$, 73, 6)
              Call REPLA(Y$, MKD$(CDbl(CAPENDE)), 79, 8)
              Call REPLA(Y$, MKD$(CDbl(CAPENDE)), 95, 8)
              Call REPLA(Y$, MKS$(60 * HORAPREP), 103, 4)
              Call REPLA(Y$, MKS$(60 * HORAPROD), 111, 4)
              Call REPLA(Y$, MKS$(60 * HORAPREP + 60 * HORAPROD), 131, 4)
              Call REPLA(Y$, EQAUX$, 203, 18)
              '
              Call REGAPP("OPERFAB", Y$)
              HAYSEC% = 1
              '
         End If
         Next TTL&
         '
         If HAYSEC% < 1 Then
             Call COMUNI("No hay Operaciones Standard\para el Codigo " + TRIM$(CODEXT$(CI%)) + "\" + TRIM$(DESCRIT$(CI%)))
         End If
         '
      Next I&
      '
      TTS$ = REGBLAN$("ORFANUE")
      Call REPLA(TTS$, MKS$(NROPED&), 1, 4)
          If QQ% > 1 Then REFE$ = "VARIOS": CI% = 0: UNI$ = "UNID": CAPENDE = 0
      Call REPLA(TTS$, REFE$, 5, 64)
      Call REPLA(TTS$, MKI$(FEPED%), 69, 2)
      Call REPLA(TTS$, MKI$(CI%), 71, 2)
      Call REPLA(TTS$, UNI$, 73, 4)
      Call REPLA(TTS$, MKD$(CDbl(CAPENDE)), 89, 8)
      Call REPLA(TTS$, DESTI$, 101, 8)
      Call REPLA(TTS$, MKI$(FEPED%), 109, 2)
      Call REPLA(TTS$, Chr$(PRIORI%), 111, 1)
      Call REPLA(TTS$, Chr$(STATI%), 112, 1)
      Call REPLA(TTS$, MKI$(FEPED%), 113, 2)
      '
      Call REGAPP("ORFANUE", TTS$)
      '
      INI& = 1 + ULTREG&("ORDEFABR")
      FIN& = ULTREG&("ORDEFABR")
      If ULTREG&("ECORDEF") > FIN& Then FIN& = ULTREG&("ECORDEF")
      '
      For I& = INI& To FIN&
         Call GRAREG("ORDEFABR", String$(256, 0), I&)
         Call GRAREG("ECORDEF", String$(64, 0), I&)
      Next I&
      '
      Call CIERRARCH("ORDEFABR")
      Call CIERRARCH("ECORDEF")
      Call CIERRARCH("ECORDEP")
      Call CIERRARCH("ORFANUE")
      '
      Screen.MousePointer = 1
      GoTo 1000
      '
PRIPEDNOVO:
    '
    TTXX$ = Mid$(PEDENCA$, 169, 32): Call GRAREG("!OBSERVOF", TTXX$, 1)
    TTXX$ = Mid$(PEDENCA$, 201, 32): Call GRAREG("!OBSERVOF", TTXX$, 2)
    TTXX$ = Mid$(PEDENCA$, 233, 32): Call GRAREG("!OBSERVOF", TTXX$, 3)
    Call SETULTREG("!OBSERVOF", 3)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = CONTROL$("", "FORPECLI")
    If FORIPRE$ = "" Then Return
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = HOS$
    NROPE& = CVS(Left$(PEDENCA$, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    TIPODOC$ = "Pedido de Cliente"
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "REF-MAT1"
    CODTABU1$(5) = "COLOR"
    CODTABU1$(6) = "UNIMED"
    CACOLTAB1% = 6
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    For U& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", U&)
       CIXI% = CVI(Left$(X$, 2))
       JJX$ = CODEXT$(CIXI%)
       DECO$ = DESCRIT$(CIXI%)
       FENTX$ = UNIMED$(CIXI%)
       CAPEX$ = CSTRING$(Mid$(X$, 47, 4), 3, 10, 1, 1)
       LG$ = "": If CVS(Mid$(X$, 27, 4)) > 1 Then LG$ = TRIM$(Str$(CVS(Mid$(X$, 27, 4))))
       AN$ = "": If CVS(Mid$(X$, 31, 4)) > 1 Then AN$ = TRIM$(Str$(CVS(Mid$(X$, 31, 4))))
       EP$ = "": If CVS(Mid$(X$, 35, 4)) > 1 Then EP$ = TRIM$(Str$(CVS(Mid$(X$, 35, 4))))
       MEDI$ = "": If LG$ + AN$ + EP$ <> "" Then MEDI$ = LG$ + " x " + AN$ + " x " + EP$
       MEDI$ = TRIM$(MEDI$)
       ITOTX$ = Mid$(X$, 3, 16)
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = JJX$
       TETABU1$(2, CAITAB1%) = DECO$
       TETABU1$(3, CAITAB1%) = CAPEX$
       TETABU1$(4, CAITAB1%) = MEDI$
       TETABU1$(5, CAITAB1%) = ITOTX$
       TETABU1$(6, CAITAB1%) = FENTX$
       '
240 Next U&
    '
    CODPARAM$(1) = "NRO-OCOM"
    DOCPARAM$(1) = TRIM$(Mid$(PEDENCA$, 57, 16))
    CODPARAM$(2) = "DESTINO"
    DOCPARAM$(2) = TRIM$(Mid$(PEDENCA$, 73, 32)) + " - " + TRIM$(Mid$(PEDENCA$, 105, 32))
    CODPARAM$(3) = "TRANSPOR"
    DOCPARAM$(3) = TRIM$(Mid$(PEDENCA$, 137, 32))
    CODPARAM$(4) = "FECH-VEN"
    DOCPARAM$(4) = FECHATEX$(HOII%)
    CODPARAM$(5) = "LIST-PRE"
    LPRE% = Asc(Mid$(PEDENCA$, 268, 1))
    DOCPARAM$(5) = TRIM$(Str$(LPRE%)) + " - " + DELIPRE$(LPRE%)
    CODPARAM$(6) = "COND-PAG"
    CPAG% = Asc(Mid$(PEDENCA$, 269, 1))
    DOCPARAM$(6) = TRIM$(Str$(CPAG%)) + " - " + DECONPA$(CPAG%)
    CODPARAM$(7) = "VENDEDOR"
    VENDX% = Asc(Mid$(PEDENCA$, 270, 1))
    DOCPARAM$(7) = TRIM$(Str$(VENDX%)) + " - " + DEVENDE$(VENDX%)
    CAPARDOC% = 7
    '
    Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
    '
Return
     '
End Sub


Sub PRIPEDNOVO()

End Sub
