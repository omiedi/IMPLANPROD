VERSION 5.00
Begin VB.Form FAPEDMAI 
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
   Icon            =   "FAPEDMAI.frx":0000
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
      Picture         =   "FAPEDMAI.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   315
      Width           =   750
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
               Picture         =   "FAPEDMAI.frx":0316
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
               Picture         =   "FAPEDMAI.frx":0758
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
               Picture         =   "FAPEDMAI.frx":0A62
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
               Picture         =   "FAPEDMAI.frx":0D6C
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
               Picture         =   "FAPEDMAI.frx":1076
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
      Picture         =   "FAPEDMAI.frx":1380
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
      Picture         =   "FAPEDMAI.frx":17C2
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
      Picture         =   "FAPEDMAI.frx":1ACC
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
      Picture         =   "FAPEDMAI.frx":1DD6
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
         Left            =   1180
         ScaleHeight     =   585
         ScaleWidth      =   2220
         TabIndex        =   33
         Top             =   525
         Width           =   2280
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
            Picture         =   "FAPEDMAI.frx":1F20
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Repetir Impresion de Pedido Registrado"
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
            Left            =   1660
            Picture         =   "FAPEDMAI.frx":258A
            Style           =   1  'Graphical
            TabIndex        =   36
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
            Picture         =   "FAPEDMAI.frx":26D4
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Registrar Pedido de Cliente"
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
            Left            =   1105
            Picture         =   "FAPEDMAI.frx":2B16
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
      Height          =   3690
      Left            =   5145
      TabIndex        =   1
      Top             =   1470
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   120
         Left            =   420
         ScaleHeight     =   60
         ScaleWidth      =   3735
         TabIndex        =   51
         Top             =   3360
         Width           =   3795
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   37
         Top             =   1575
         Width           =   2535
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
            Left            =   0
            Picture         =   "FAPEDMAI.frx":2E20
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Pedidos Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Código de Artículo o Material"
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
            TabIndex        =   39
            Top             =   105
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   48
         Top             =   1050
         Width           =   2535
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
            Left            =   0
            Picture         =   "FAPEDMAI.frx":312A
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Pedidos Pendientes por Cliente"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Número de Cliente"
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
            TabIndex        =   49
            Top             =   70
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   45
         Top             =   525
         Width           =   2535
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
            Left            =   0
            Picture         =   "FAPEDMAI.frx":356C
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Detalle por Número de Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencial por Número de Pedido"
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
            TabIndex        =   46
            Top             =   70
            Width           =   1695
         End
      End
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
         TabIndex        =   40
         Top             =   2520
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
            TabIndex        =   42
            Top             =   270
            Value           =   -1  'True
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
            TabIndex        =   41
            Top             =   270
            Width           =   1065
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
      Picture         =   "FAPEDMAI.frx":3876
      Stretch         =   -1  'True
      Top             =   5355
      Width           =   10920
   End
End
Attribute VB_Name = "FAPEDMAI"
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
    If InStr(UCase$(EMPRELI$), "SABO") < 1 Then
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
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
       Call OPNOIN("Revertir Anulación de Pedido Itemizado")
       GoTo 99
    End If
    Call DESANUPED
99  Command43.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call FINAL("*CPROGEN")
    Command25.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ADMIPED")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("FORMULAS", "Fórmulas y Estructuras de Producto")
    Command3.Enabled = True
End Sub

Private Sub Command41_Click()
    Command41.Enabled = False
    If FORCAPED% < 1 Or FORCAPED% > 2 Then
       FORCAPED% = 1
       If Control$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then FORCAPED% = 2
    End If
    '
    If FORCAPED% = 2 Then
         FORCARPED.Show
       Else
         Call PEDCLINUE
    End If
    '
    Command41.Enabled = True
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
    Call LIPEDIT(4)
    Command64.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ADMIPED")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    ORDESPACH.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Command45.Enabled = False
    Call ANULAPEDMAI
99  Command45.Enabled = True
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
    Call LIPEDIT(3)
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    Call LIPEDIT(2)
    Command63.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub
