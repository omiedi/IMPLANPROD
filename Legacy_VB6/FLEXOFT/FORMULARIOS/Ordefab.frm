VERSION 5.00
Begin VB.Form ORDEFAB 
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O/F's - Ordenes de Fabricación"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9840
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Ordefab.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Operaciones Varias"
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
      Left            =   210
      TabIndex        =   82
      Top             =   5355
      Width           =   4635
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   93
         Top             =   420
         Width           =   1695
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
            Picture         =   "Ordefab.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Remito de Materiales en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Repara- ciones"
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
            Left            =   735
            TabIndex        =   95
            Top             =   105
            Width           =   1110
         End
      End
      Begin VB.Frame Frame10 
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
         TabIndex        =   86
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command48 
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
            TabIndex        =   92
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command47 
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
            TabIndex        =   91
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command39 
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
            TabIndex        =   90
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command38 
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
            TabIndex        =   89
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command37 
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
            TabIndex        =   88
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command32 
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
            TabIndex        =   87
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   83
         Top             =   420
         Width           =   1800
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
            Left            =   0
            Picture         =   "Ordefab.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Remito de Materiales en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Servicios  de Terceros"
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
            Left            =   630
            TabIndex        =   85
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Command13"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5145
      TabIndex        =   72
      Top             =   1260
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command12 
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
      Height          =   645
      Left            =   5040
      Picture         =   "Ordefab.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Control de la Base de Datos"
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5040
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   69
      Top             =   1050
      Width           =   765
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Asignación de Materias Primas"
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
      Left            =   210
      TabIndex        =   31
      Top             =   3990
      Width           =   4635
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1665
         TabIndex        =   66
         Top             =   420
         Width           =   1730
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
            Left            =   0
            Picture         =   "Ordefab.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Transferencia Materiales de Depósito a Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Vale de Materiales"
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
            Left            =   660
            TabIndex        =   68
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   39
         Top             =   420
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
            Picture         =   "Ordefab.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Control y Gestión de Reservas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales Reservados"
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
            Left            =   630
            TabIndex        =   41
            Top             =   105
            Width           =   1320
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
         TabIndex        =   32
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0F8&
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
      Height          =   2325
      Left            =   210
      TabIndex        =   16
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   210
      Width           =   4635
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   73
         Top             =   1470
         Width           =   2535
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
            Left            =   0
            Picture         =   "Ordefab.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Base de Datos de Proveedores y Subcontratistas"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedores y Subcontratistas"
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
            TabIndex        =   74
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   75
         Top             =   945
         Width           =   2535
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
            Left            =   0
            Picture         =   "Ordefab.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Padrón de Clien- tes y Deudores"
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
            TabIndex        =   76
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   77
         Top             =   420
         Width           =   2535
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
            Picture         =   "Ordefab.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Maestro de Artí- culos de Stock"
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
            TabIndex        =   78
            Top             =   50
            Width           =   1800
         End
      End
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
         TabIndex        =   17
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   1890
            Width           =   1590
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
      Left            =   5985
      Picture         =   "Ordefab.frx":1AD2
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   315
      Width           =   855
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
      Left            =   6930
      Picture         =   "Ordefab.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programación de Carga"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   1575
      Width           =   4635
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   60
         Top             =   525
         Width           =   1800
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
            Picture         =   "Ordefab.frx":221E
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Consulta de Tiempo Total Programado por Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Total por Equipo"
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
            Left            =   630
            TabIndex        =   62
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   57
         Top             =   525
         Width           =   1800
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
            Picture         =   "Ordefab.frx":2528
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Consulta y Actualización de Operaciones Proyectadas por O.F."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones Programadas"
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
            Left            =   630
            TabIndex        =   59
            Top             =   105
            Width           =   1320
         End
      End
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
      Left            =   7875
      Picture         =   "Ordefab.frx":2672
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   315
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
      Left            =   8820
      Picture         =   "Ordefab.frx":297C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Base de Datos de O-F's"
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
      Left            =   210
      TabIndex        =   3
      Top             =   2625
      Width           =   4635
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   440
         ScaleHeight     =   585
         ScaleWidth      =   3675
         TabIndex        =   24
         Top             =   420
         Width           =   3730
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
            Left            =   3150
            Picture         =   "Ordefab.frx":2AC6
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Ver y Re-Imprimir Copia de Orden de Fabricación Emitida"
            Top             =   0
            Width           =   530
         End
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
            Left            =   3675
            Picture         =   "Ordefab.frx":2DD0
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Orden de Disposicion de Tintorería"
            Top             =   0
            Width           =   530
         End
         Begin VB.CommandButton Command44 
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
            Left            =   1580
            Picture         =   "Ordefab.frx":2F1A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Suspensión de Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.CommandButton Command42 
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
            Left            =   520
            Picture         =   "Ordefab.frx":3224
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Lanzamiento de Orden de Fabricación"
            Top             =   0
            Width           =   530
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
            Left            =   2620
            Picture         =   "Ordefab.frx":352E
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Modificación de Prioridad y Sector Productivo"
            Top             =   0
            Width           =   530
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
            Picture         =   "Ordefab.frx":3678
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Generar Orden de Fabricación"
            Top             =   0
            Width           =   530
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
            Left            =   1050
            Picture         =   "Ordefab.frx":3982
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Cierre de Orden de Fabricación"
            Top             =   0
            Width           =   530
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
            Left            =   2100
            Picture         =   "Ordefab.frx":3C8C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anulación de Orden de Fabricación"
            Top             =   0
            Width           =   530
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
         TabIndex        =   4
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0F8&
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
      Left            =   5040
      TabIndex        =   2
      Top             =   3045
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0F8&
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
         TabIndex        =   54
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0F8&
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
            TabIndex        =   56
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0F8&
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
            TabIndex        =   55
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
         TabIndex        =   48
         Top             =   525
         Width           =   3795
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
            Left            =   1365
            Picture         =   "Ordefab.frx":3F96
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Pendientes por Sector Productivo"
            Top             =   0
            Width           =   600
         End
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
            Picture         =   "Ordefab.frx":42A0
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Pendientes por Número de Orden de Fabricación"
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
            Picture         =   "Ordefab.frx":45AA
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Consulta General de la Base de Datos"
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
            Picture         =   "Ordefab.frx":48B4
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Consulta / Listado Selectivo por Estado"
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
            Picture         =   "Ordefab.frx":49FE
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Consulta / Listado Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B-D de O/F's "
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
            TabIndex        =   52
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F8&
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
      Height          =   1380
      Left            =   5040
      TabIndex        =   1
      Top             =   5250
      Width           =   4635
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   45
         Top             =   525
         Width           =   1800
         Begin VB.CommandButton Command51 
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
            Picture         =   "Ordefab.frx":4D08
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Acceso Directo a Informes de Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
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
            Height          =   645
            Left            =   630
            TabIndex        =   47
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   42
         Top             =   525
         Width           =   1800
         Begin VB.CommandButton Command52 
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
            Picture         =   "Ordefab.frx":514A
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Acceso Directo a Programación de Carga de Máquinas (MRP-II)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga de Máquinas"
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
            Left            =   735
            TabIndex        =   44
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9870
      Y1              =   6825
      Y2              =   6825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Modular Integrado de Procesamiento de Datos Línea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3780
      TabIndex        =   0
      Top             =   7035
      Width           =   5895
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "Ordefab.frx":5454
      Stretch         =   -1  'True
      Top             =   6825
      Width           =   10920
   End
End
Attribute VB_Name = "ORDEFAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C05", "Consulta O/F's por Sector")
        Else
           ZZ$ = REGSEL$("SECPROD")
           If Len(ZZ$) <= 4 Then GoTo 99
           SESEL$ = Left$(ZZ$, 4)
           J& = 0
           Screen.MousePointer = 11
           For U& = 1 To ULTREG&("ORDEFABR")
               X$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(X$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(X$, 89, 8))
                     CAPRA# = CVD(Mid$(X$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     If CAPEN# > 0 Then
10                      SPX$ = Mid$(X$, 113, 4)
                        If ECOARCH$("SECPROD", SPX$) = "" Then
                          CIXI% = CVI(Mid$(X$, 83, 2))
                          OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + String$(4, 0))
                          If OBX$ = "" Then GoTo 99
                          SPX$ = Mid$(OBX$, 3, 4)
                          Call REPLA(X$, SPX$, 113, 4)
                          Call GRAREG("ORDEFABR", X$, U&)
                          GoTo 10
                        End If
15                      If SPX$ = SESEL$ Then
                           Call REPLA(X$, MKD$(CAPEN#), 159, 8)
                           Call REPLA(X$, MKD$(0), 167, 8)
                           J& = J& + 1
                           Call GRAREG("ORFASEL", X$, J&)
                        End If
                     End If
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", J&)
           Call CIERRARCH("ORFASEL")
           Call CIERRARCH("*.*")
           Call HEADERS("ORFASEL", "")
           Call HEADERS("ORFASEL", "Sector: " + ECOARCH$("SECPROD", SESEL$))
           Screen.MousePointer = 1
           '
           Call FINAL("*SORDEFA")
    End If
99  Command10.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call VALEMATER
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If COMALTER%("Control y Recuperación de la Base de Datos de O/F's\\Cancelar\Realizar Control") = 2 Then
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("ORDEFABR")
         XX$ = REGLEIDO$("ORDEFABR", U&)
         If InStr(XX$, "OF-") > 69 Then
            While InStr(XX$, "OF-") > 69
               XX$ = Mid$(XX$, 2)
            Wend
            If Abs(CVS(Left$(XX$, 4)) - Val(Mid$(XX$, 72, 6))) > 0.05 Then
               Call REPLA(XX$, MKS$(Val(Mid$(XX$, 72, 6))), 1, 4)
            End If
            Call GRAREG("ORDEFABR", XX$, U&)
         End If
      Next U&
      Call CIERRARCH("ORDEFABR")
      Call ACTECOR(0)
      Screen.MousePointer = 1
   End If
   Command12.Enabled = True
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   OPTEROPER.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    OPORFABRI.Show 1
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call CONECRUN("AMACLI01", "Padrón Maestro de Clientes")
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Padrón de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("REPARA02", "Reparaciones")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    DISPOTINT.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command40_Click()
Dim ORFA$(2048), PUNCOD$(2048)
For U& = 1 To ULTREG&("RESERVA")
XX$ = REGLEIDO$("RESERVA", U&)
CII% = CVI(Left$(XX$, 2))
NOFAI$ = Mid$(XX$, 3, 12)
For KKI% = 1 To KKJ%
If ORFA$(KKI%) = NOFAI$ Then
   PUNCOD$(KKI%) = PUNCOD$(KKI%) + MKI$(CII%)
   GoTo 10
End If
Next KKI%
KKJ% = KKJ% + 1
ORFA$(KKJ%) = NOFAI$
PUNCOD$(KKJ%) = MKI$(CII%)
10 Next U&
'
For KKI% = 1 To KKJ%
PUCO$ = PUNCOD$(KKI%)
For I1% = 1 To Len(PUCO$) Step 2
CI1% = CVI(Mid$(PUCO$, I1%, 2))
For I2% = I1% + 2 To Len(PUCO$) Step 2
CI2% = CVI(Mid$(PUCO$, I2%, 2))
If CI1% = CI2% Then
MsgBox ORFA$(KKI%)
End If
Next I2%
Next I1%
Next KKI%
Stop
'
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    Call ORFPRI
    Command43.Enabled = True
End Sub

Private Sub Command44_Click()
    Command44.Enabled = False
    Call ORFSUS
    Command44.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call FINAL("*CPROGEN")
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
    Command26.Enabled = False
    Call CONECRUN("#AUTENTR", "")
    Command26.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    Call FINAL("*LPLAFAB")
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub

Private Sub Command30_Click()
    Command30.Enabled = False
    Call FINAL("*LPLACOM")
    Command30.Enabled = True
End Sub

Private Sub Command42_Click()
    Command42.Enabled = False
    Call ORFLAN
    Command42.Enabled = True
End Sub

Private Sub Command41_Click()
    If EXISTE%("ADMIOTI.EXE") > 0 Then
        TXT$ = "La Presente Versión de FLEXOFT\"
        TXT$ = TXT$ + "sólo Permite Nuevas Ordenes de\"
        TXT$ = TXT$ + "Fabricación por Generación de la\"
        TXT$ = TXT$ + "Carga de Máquina a Partir de los\"
        TXT$ = TXT$ + "Pedidos Externos o Internos."
        TXT$ = TXT$ + "\\Generar Carga\Cancelar"
        If COMALTER%(TXT$) = 1 Then
            Call CONECRUN("#ADMIOTI", "")
        End If
        Exit Sub
    End If
    Command41.Enabled = False
    Call ORFNUE(1)
    Command41.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   OPX% = COMALTER%("Opciones Disponibles:\\Ver y Re-Imprimir\Re-Generar")
   If OPX% = 1 Then
        Call VERORFA
      ElseIf OPX% = 2 Then
        Call REPRINORFA
   End If
   Command5.Enabled = True
End Sub

Private Sub Command51_Click()
    Command51.Enabled = False
    Call CONECRUN("INFOFAB", "Informes de Producción")
    Command51.Enabled = True
End Sub

Private Sub Command52_Click()
    Command52.Enabled = False
    Call CONECRUN("CARGMAQ", "Carga de Máquinas")
    Command52.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call TOTEQUIP
'    Call FINAL("*TOTIEPRO")
    Command6.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C02", "Consulta O/F's por Producto")
        Else
           J& = 0
           Screen.MousePointer = 11
           For U& = 1 To ULTREG&("ORDEFABR")
               X$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(X$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(X$, 89, 8))
                     CAPRA# = CVD(Mid$(X$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     Call REPLA(X$, MKD$(CAPEN#), 159, 8)
                     Call REPLA(X$, MKD$(0), 167, 8)
                     J& = J& + 1
                     Call GRAREG("ORFASEL", X$, J&)
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", J&)
           Call CIERRARCH("*.*")
           Screen.MousePointer = 1
           '
           Call FINAL("*SORDEFA")
    End If
    Command64.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call ACTECOR(1)
'    Call FINAL("*OPERFAB")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ORDEFABR")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    OPRESERVA.Show
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Call ORFANU
End Sub

Private Sub Command46_Click()
    Call ORFCIE
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    Call REVIORFLAN
    If Option1.Value = True Then
           Call CONECRUN("*CONORFA", "Consulta Base de datos de O.F's")
        Else
           Call OPNOIN("")
    End If
    Command61.Enabled = True
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
    Call GRAREG("ORFASTA", Chr$(0) + "Proyectadas", 1)
    Call GRAREG("ORFASTA", Chr$(1) + "Lanzadas", 2)
    Call GRAREG("ORFASTA", Chr$(2) + "Suspendidas", 3)
    Call GRAREG("ORFASTA", Chr$(3) + "Cerradas", 4)
    Call GRAREG("ORFASTA", Chr$(9) + "Anuladas", 5)
    Call CIERRARCH("ORFASTA")
    If Option1.Value = True Then
           Call CONECRUN("*CONORFS", "Consulta O.F's por Estado")
        Else
           OPX$ = OBJPLA$("SELORFSTA", "")
           If OPX$ <> "" Then
              If ECOARCH$("ORFASTA", OPX$) <> "" Then
                 J& = 0
                 Screen.MousePointer = 11
                 Call HEADERS("ORFASEL", "")
                 Call HEADERS("ORFASEL", "Estado: " + ECOARCH$("ORFASTA", OPX$))
                 For U& = 1 To ULTREG&("ORDEFABR")
                    X$ = REGLEIDO$("ORDEFABR", U&)
                    If Mid$(X$, 108, 1) = OPX$ Then
                       J& = J& + 1
                       Call GRAREG("ORFASEL", X$, J&)
                    End If
                Next U&
                Call SETULTREG("ORFASEL", J&)
                Call CIERRARCH("*.*")
                Screen.MousePointer = 1
                '
                Call FINAL("*LIORFST")
                '
             End If
          End If
    End If
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C01", "Consulta por Número de O.F.")
        Else
           Call OPNOIN("")
    End If
    Command63.Enabled = True
End Sub

Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If InStr(EPAC$, "BOREN") > 0 Then
       Command12.Visible = True
    End If
    '
    Call VERIDEST
    If InStr(UCase$(EPAC$), "BORENS") > 0 Then
       Picture1.Width = Picture1.Width + Command4.Width
       Picture1.Left = Picture1.Left - Command4.Width / 2
    End If
    '
    If EXISTE%("REPARA02.EXE") < 1 Then
      Command3.Enabled = False
      Label5.Enabled = False
    End If
    '
    Show
    Call REVIBOPER
    '
End Sub
'

Sub VALEMATER()
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   MODOVA% = ALTERNA%("Vale de Preparacion\Vale de Transferencia")
   If MODOVA% < 1 Or MODOVA% > 2 Then Exit Sub
   '
   If MODOVA% = 2 Then
     If DERACCE%("VALETRAN", "Vale de Transferencia de Stocks") < 2 Then
       Exit Sub
     End If
   End If
   '
10 CAOSEL% = 0
   XX$ = REGSEL$("SECPROD")
   If Len(XX$) <= 4 Then
      Exit Sub
   End If
   '
   SECCI = CVS(Left$(XX$, 4))
   SECPROD$ = Mid$(XX$, 5, 32)
   DEPODEST% = Asc(Mid$(XX$, 81, 1))
   '
   DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
   EPAC$ = EMPREAC$
   If InStr(EPAC$, "DOSIVAC") > 0 Then
      Call SELECHO("TADEPOSI")
      DEPORIG% = Val(VALACT1$("TADEPOSI"))
      If DEPORIG% < 1 Then Exit Sub
   End If
   If DEPORIG% = DEPODEST% Then
      Call MENSERR(24, "Operación Incorrecta.\Deposito Origen igual a Destino.")
      GoTo 10
   End If
   DEPOSI$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPORIG%)))
   Call BLANARCH("!VALEMAT")
   '
   Screen.MousePointer = 11
   For U& = 1 To NUMAS%
      CANECE(U&) = 0
   Next U&
   '
   'SELECCIONAR ORDENES LANZADAS DE ESE SECTOR
11 FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         If CVS(Mid$(XX$, 113, 4)) = SECCI Then
            CAOSEL% = CAOSEL% + 1
            ORFASEL$(CAOSEL%) = Mid$(XX$, 69, 12)
         End If
      End If
   Next U&
   '
   'CARGAR MATERIALES RESERVADOS PARA O.F'S
   FIN& = ULTREG&("RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("RESERVA", U&)
      ORFA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) - CVD(Mid$(XX$, 27, 8))
      If CANTI < 0 Then CANTI = 0
      '
      If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
         CANTI = CANTI / 1000
      End If
      '
      For KK% = 1 To CAOSEL%
         If ORFA$ = ORFASEL$(KK%) Then GoTo 15
      Next KK%
      GoTo 19
      '
15    If CI% > 0 Then
         If CI% <= NUMAS% Then
            CANECE(CI%) = CANECE(CI%) + CANTI
         End If
      End If
      '
19 Next U&
   '
   FIN& = NUINV%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      II$ = REGLEIDO$("INVERT", U& + 1)
      CI% = CVI(Mid$(II$, 17, 2))
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            If CANECE(CI%) > 0 Then
               If Abs(SEPRODUC(CI%) - SECCI) > 0.0005 Then
                  STODEST = STODEPOS(CI%, DEPODEST%)
                  STOCENT = STODEPOS(CI%, DEPORIG%)
                  NENETA = CANECE(CI%) - STODEST
                  If NENETA < 0 Then NENETA = 0
                  If NENETA > 0.05 Then
                     VALEMATX = NENETA
                     If VALEMATX > STOCENT Then VALEMATX = STOCENT
                     If VALEMATX < 0 Then VALEMATX = 0
                     '
                     ZZ$ = REGBLAN$("VALEMAT")
                     Call REPLA(ZZ$, MKI$(CI%), 1, 2)
                     Call REPLA(ZZ$, MKS$(CANECE(CI%)), 3, 4)
                     Call REPLA(ZZ$, MKS$(STODEST), 7, 4)
                     Call REPLA(ZZ$, MKS$(NENETA), 11, 4)
                     Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)
                     Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)
                     Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
                     Call REPLA(ZZ$, "......", 27, 6)
                     Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
                     Call REGAPP("!VALEMAT", ZZ$)
                  End If
               End If
            End If
         End If
      End If
   Next U&
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1
   '
   If MODOVA% = 1 Then
      TITUPA$ = "Vale de PREPARACION" + " - de " + DEPOSI$ + " a " + SECPROD$
      VTB% = VENTABU%("VALEMAP" + "/TITUPAN=" + TITUPA$)
      If VTB% < 0 Then GoTo 10
      '
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To ULTREG&("!VALEMAT")
         ZZ$ = REGLEIDO$("!VALEMAT", U&)
         JJ& = U&
         Call GRAREG("VALEMAT", ZZ$, JJ&)
      Next U&
      Call CIERRARCH("!VALEMAT")
      Call SETULTREG("VALEMAT", JJ&)
      Call CIERRARCH("VALEMAT")
      Call HEADERS("VALEMAT", "")
      SECCX$ = TRIM$(CSTRING$(MKS$(SECCI), 3, 8, 3, 2))
      Call HEADERS("VALEMAT", "Seccion: " + SECCX$ + " - " + SECPROD$)
      Screen.MousePointer = 1
      Call CONECRUN("*PRIVALMA", "Vale de Preparacion de Materiales")
      GoTo 10
   End If
   '
   TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + DEPOSI$ + " a " + SECPROD$
   VTB% = VENTABU%("VALEMAT" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then GoTo 10
   '
   REGXX& = ULTREG&("TRANSFE")
   NRO& = 1
   If REGXX& > 0 Then
       NRO& = 1 + CVS(Mid$(REGLEIDO$("TRANSFE", REGXX&), 45, 4))
   End If
   If NRO& <= REGXX& Then
       Screen.MousePointer = 11
       For U& = 1 To REGXX&
           X0$ = REGLEIDO$("TRANSFE", U&)
           If CVS(Mid$(X0$, 45, 4)) >= NRO& Then
               NRO& = 1 + CVS(Mid$(X0$, 45, 4))
           End If
       Next U&
       Screen.MousePointer = 1
   End If
   '
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("TRANSFE")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
                             While Len(DOSEC$) < 10
                               DOSEC$ = DOSEC$ + " "
                             Wend
                              '
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + SECPROD$
         '
         CANTI = (-1) * CAN
         Call MOVISTO(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPORIG%, CANTI)
         '
         REFE$ = "TR.de " + DEPOSI$
         CANTI = CAN
         Call MOVISTO(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPODEST%, CANTI)
         '
         XXX$ = REGBLAN$("TRANSFE")
         Call REPLA(XXX$, MKI$(FECHA%), 1, 2)
         Call REPLA(XXX$, Chr$(DEPORIG%), 3, 1)
         Call REPLA(XXX$, Chr$(DEPODEST%), 4, 1)
         Call REPLA(XXX$, MKI$(CI1%), 5, 2)
         Call REPLA(XXX$, MKI$(CI1%), 7, 2)
         Call REPLA(XXX$, LOTE$, 9, 16)
         Call REPLA(XXX$, LOTE$, 25, 16)
         Call REPLA(XXX$, MKS$(CAN), 41, 4)
         Call REPLA(XXX$, MKS$(NRO&), 45, 4)
         Call REGAPP("TRANSFE", XXX$)
         '
      End If
   Next U&
   '
   Call CIERRARCH("TRANSFE")
   Call CIERRARCH("MOVISTO")
   Call BAJALDISCO
   '
   FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DEPORIG%)) + " - " + DEPOSI$
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = CSTRING$(MKS$(SECCI), 3, 8, 3, 2) + " - " + SECPROD$
      CAPARDOC% = 2
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CACOLTAB1% = 5
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For U& = 1 To ULTREG&("!VALEMAT")
         XX$ = REGLEIDO$("!VALEMAT", U&)
         If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
            '
            CICRU% = CVI(Mid$(XX$, 1, 2))
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            KICOL1$ = Str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = UNIMED$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            '
         End If
      Next U&
      '
      If CAITAB1% < 1 Then
         Screen.MousePointer = 1
         GoTo 10
      End If
      '
      Call PRINTDOC("FORVALMA")   ' Vale de Materiales
      Call BAJALDISCO
      '
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub

Sub VERORFA()
    '
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    TIDOCUM$ = "Orden de Fabricación " + NORFA$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  End Sub

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
