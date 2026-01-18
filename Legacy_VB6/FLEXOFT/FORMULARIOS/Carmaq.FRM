VERSION 5.00
Begin VB.Form CARMAQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.R.P. - Programación de Carga de Máquinas"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10860
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Form"
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
      Left            =   9840
      Picture         =   "Carmaq.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Configuración Formulario Planilla de Carga"
      Top             =   4410
      Width           =   855
   End
   Begin VB.Frame Frame1 
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
      Height          =   3585
      Left            =   210
      TabIndex        =   0
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Productos y Rutas"
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
         TabIndex        =   2
         Top             =   525
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   31
            Top             =   420
            Width           =   3165
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
               Left            =   2520
               Picture         =   "Carmaq.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Hojas de Ruta - Secuencia de Operaciones de Fabricación"
               Top             =   0
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
               Left            =   0
               Picture         =   "Carmaq.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Archivos Maestros de Artículos de Stock"
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
               TabIndex        =   33
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame8 
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
         TabIndex        =   11
         Top             =   2100
         Width           =   4005
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   35
            Top             =   420
            Width           =   3165
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
               Picture         =   "Carmaq.frx":0A56
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Base de Datos de Proveedores y Subcontratistas"
               Top             =   0
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
               Left            =   -20
               Picture         =   "Carmaq.frx":0D60
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
               Top             =   0
               Width           =   560
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
               TabIndex        =   37
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame5 
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
         Begin VB.CommandButton Command13 
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
            TabIndex        =   9
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
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
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   7
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   5
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command4 
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
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.CommandButton Command34 
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
      Left            =   9840
      Picture         =   "Carmaq.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   2205
      Width           =   855
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
      Left            =   9840
      Picture         =   "Carmaq.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
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
      Left            =   9840
      Picture         =   "Carmaq.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Complementos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   210
      TabIndex        =   12
      Top             =   3990
      Width           =   4635
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1680
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   39
         Top             =   1785
         Width           =   2325
         Begin VB.CommandButton Command35 
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
            Picture         =   "Carmaq.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Acceso Directo a Calculo de Necesidades de Materiales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Programación de Materiales MRP-II"
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
            TabIndex        =   41
            Top             =   105
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   48
         Top             =   1365
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
            Picture         =   "Carmaq.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Acceso Directo a Gestión de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Compras"
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
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   42
         Top             =   945
         Width           =   2325
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
            Picture         =   "Carmaq.frx":1F14
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Acceso Directo a Informes de Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes de Producción"
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
            TabIndex        =   44
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   45
         Top             =   525
         Width           =   2325
         Begin VB.CommandButton Command24 
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
            Picture         =   "Carmaq.frx":221E
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Acceso Directo a Ordenes de Fabricación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenes de Fabricación"
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
            TabIndex        =   47
            Top             =   0
            Width           =   1695
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
         TabIndex        =   13
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
      Height          =   5055
      Left            =   5040
      TabIndex        =   10
      Top             =   1680
      Width           =   4635
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Operaciones Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   315
         TabIndex        =   23
         Top             =   420
         Width           =   4005
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Formato"
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
            Left            =   420
            TabIndex        =   55
            Top             =   1260
            Width           =   3165
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Tabulado"
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
               Left            =   210
               TabIndex        =   57
               Top             =   315
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Gr.GANTT"
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
               Left            =   1785
               TabIndex        =   56
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   51
            Top             =   420
            Width           =   3165
            Begin VB.CommandButton Command26 
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
               Picture         =   "Carmaq.frx":2528
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Operaciones Pendientes por Orden de Fabricación"
               Top             =   0
               Width           =   560
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
               Left            =   2550
               Picture         =   "Carmaq.frx":2832
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Operaciones Pendientes por Máquina o Equipo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Secuencia de Ope- raciones Pendientes"
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
               TabIndex        =   53
               Top             =   105
               Width           =   1800
            End
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listados Varios de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   315
         TabIndex        =   22
         Top             =   2625
         Width           =   4005
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1365
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   58
            Top             =   1380
            Width           =   2325
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
               Left            =   0
               Picture         =   "Carmaq.frx":2B3C
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Plan de Fechas Proyectadas de Entrega"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Fechas de Entre- ga Proyectadas"
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
               TabIndex        =   60
               Top             =   80
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   61
            Top             =   1065
            Width           =   2325
            Begin VB.CommandButton Command33 
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
               Picture         =   "Carmaq.frx":2E46
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Planilla de Horas por Sección"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Horas - Sección"
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
               TabIndex        =   63
               Top             =   60
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   64
            Top             =   735
            Width           =   2325
            Begin VB.CommandButton Command27 
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
               Picture         =   "Carmaq.frx":3150
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Planilla General de Producción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Pl. Producción"
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
               TabIndex        =   66
               Top             =   60
               Width           =   1800
            End
         End
         Begin VB.PictureBox Picture13 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   67
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command28 
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
               Picture         =   "Carmaq.frx":345A
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Planilla de Carga de Máquinas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Planilla de Carga"
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
               TabIndex        =   69
               Top             =   60
               Width           =   1695
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cálculo de Carga de Máquina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5040
      TabIndex        =   1
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   1530
         TabIndex        =   28
         Top             =   420
         Width           =   1590
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
            Picture         =   "Carmaq.frx":3764
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
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
            Height          =   540
            Left            =   735
            TabIndex        =   30
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2520
         ScaleHeight     =   585
         ScaleWidth      =   1530
         TabIndex        =   25
         Top             =   420
         Width           =   1590
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
            Left            =   -20
            Picture         =   "Carmaq.frx":3BA6
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Secuenciación de Carga de Máquina"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Re- Calcular"
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
            TabIndex        =   26
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9300
      Picture         =   "Carmaq.frx":3EB0
      Top             =   6210
      Width           =   2010
   End
End
Attribute VB_Name = "CARMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Call OPNOIN("")
End Sub

Private Sub Command13_Click()
    Call OPNOIN("")
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    If EXISTE%("ORCOMP02.EXE") > 0 Then
        Call CONECRUN("ORCOMP02", "Gestion de Ordenes de Compra")
      Else
        Call CONECRUN("OCOMPRA", "Gestion de Ordenes de Compra")
    End If
    Command14.Enabled = True
End Sub


Private Sub Command15_Click()
    Command15.Enabled = False
    If Option1.Value = True Then
         Screen.MousePointer = 11
         FIN& = ULTREG("OPERFAP")
         For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           XX$ = REGLEIDO$("OPERFAP", U&)
           TITO = 60 * CVS(Mid$(XX$, 131, 4))
           Call REPLA(XX$, MKS$(TITO), 131, 4)
           Call GRAREG("OPERFAH", XX$, U&)
         Next U&
         Call SETULTREG("OPERFAH", FIN&)
         Call CIERRARCH("OPERFAB")
         Screen.MousePointer = 1
         Call FINAL("*SECAREQ")
       ElseIf Option2.Value = True Then
         'Call OPNOIN("Grafica de GANTT Operaciones por Equipo")
         FORMUGANTT.Show 1
         YACARGANTT% = 0
    End If
    Command15.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If EXISTE%("AMACLI02.EXE") > 0 Then
        Call CONECRUN("AMACLI02", "Padrón Maestro de Clientes")
      Else
        Call CONECRUN("AMACLI", "Padrón Maestro de Clientes")
    End If
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Maestro de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If EXISTE%("AMASTO02.EXE") > 0 Then
        Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
      Else
        Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    End If
    Command2.Enabled = True
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    If EXISTE%("ORFABR02.EXE") > 0 Then
        Call CONECRUN("ORFABR02", "Ordenes de Fabricación")
      Else
        Call CONECRUN("ORDEFABR", "Ordenes de Fabricación")
    End If
    Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If EXISTE%("IFABRI02.EXE") > 0 Then
        Call CONECRUN("IFABRI02", "Informes de Producción")
      Else
        Call CONECRUN("INFOFAB", "Informes de Producción")
    End If
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
    Command26.Enabled = False
    If Option1.Value = True Then
         Call FINAL("*SECAROF")
       ElseIf Option2.Value = True Then
         FORMUGANTOF.Show 1
         YACARGANTT% = 0
    End If
    Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    Call OPNOIN("")
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
10  Call PLACARMA
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CARGMAQ")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If EXISTE%("AMASTO02.EXE") > 0 Then
        Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
      Else
        Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    End If
    Command3.Enabled = True
End Sub


Private Sub Command30_Click()
    Call OPNOIN("")
End Sub

Private Sub Command31_Click()
    Call OPNOIN("")
End Sub

Private Sub Command32_Click()
    FORPROEN.Show 1
End Sub

Private Sub Command33_Click()
    Call OPNOIN("")
End Sub

Private Sub Command34_Click()
    Call ACCDIR("CARGMAQ")
End Sub

Private Sub Command35_Click()
    Command35.Enabled = False
    If EXISTE%("PROMRP02.EXE") > 0 Then
         Call CONECRUN("PROMRP02", "Programacion de Materiales - MRP-II")
       Else
         Call CONECRUN("PROGMAT", "Programacion de Materiales - MRP-II")
    End If
    Command35.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    CONFCARG.Show
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    COMPLECALCU% = 0
    FORCAMAQ.Show 1
    DoEvents
    '
    If COMPLECALCU% = 1 Then
      Call CIERRARCH("*.*")
      Call LIBARCH("*.*")
      Call FILESORT("OPERFAP", "", 20, 21)
      Call COMUNI("Proceso Completo")
      If EXISTE%(App.EXEName + ".EXE") > 0 Then
        Call BLOQEXE("*")
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Call CONECRUN(App.EXEName, "Programación de Carga de Máquinas")
        Cancel = 0
        Close: End
      End If
    End If
    '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    OPCARMAQ.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call FINAL("*SECAREQ")
    Command9.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "CARGMAQ"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub

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

Private Sub Timer1_Timer()

End Sub
