VERSION 5.00
Begin VB.Form OFCONF04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O/F's - Ordenes de Fabricación y Envios"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   10770
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "OFCONF04.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Archivos Maestros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   180
      TabIndex        =   87
      Top             =   120
      Width           =   4955
      Begin VB.PictureBox Picture17 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   90
         Top             =   525
         Width           =   1800
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
            Height          =   600
            Left            =   0
            Picture         =   "OFCONF04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a Articulos"
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
            TabIndex        =   91
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2880
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   88
         Top             =   525
         Width           =   1800
         Begin VB.CommandButton Command49 
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
            Picture         =   "OFCONF04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a  Proveedores"
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
            TabIndex        =   89
            Top             =   105
            Width           =   1320
         End
      End
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
      Height          =   640
      Left            =   3780
      Picture         =   "OFCONF04.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   86
      ToolTipText     =   "Listado Fabricación por Rango de Fechas"
      Top             =   7245
      Width           =   650
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00008080&
      Height          =   6840
      Left            =   9870
      ScaleHeight     =   6780
      ScaleWidth      =   1005
      TabIndex        =   80
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   84
         Top             =   2280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "OFCONF04.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Left            =   120
         Picture         =   "OFCONF04.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   825
         Width           =   650
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
         Height          =   640
         Left            =   120
         Picture         =   "OFCONF04.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1590
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OFCONF04.frx":14BE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2520
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Operaciones Varias"
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
      Left            =   4830
      TabIndex        =   66
      Top             =   6195
      Width           =   4425
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   77
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
            Picture         =   "OFCONF04.frx":33E0
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Gestión de Reparaciones"
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
            TabIndex        =   79
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
         TabIndex        =   70
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   67
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
            Picture         =   "OFCONF04.frx":36EA
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Tercerizar Operaciones"
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
            TabIndex        =   69
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Command13"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4515
      TabIndex        =   59
      Top             =   7455
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asignación de Materias Primas"
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
      Left            =   0
      TabIndex        =   26
      Top             =   4680
      Width           =   4955
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2600
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   57
         Top             =   420
         Width           =   2000
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia de Materiales"
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
            TabIndex        =   58
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   340
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   34
         Top             =   420
         Width           =   2000
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
            Left            =   735
            TabIndex        =   35
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
         TabIndex        =   27
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   315
      TabIndex        =   11
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   6090
      Width           =   4425
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   60
         Top             =   525
         Width           =   1905
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
            Picture         =   "OFCONF04.frx":39F4
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Base de Datos de Proveedores y Subcontratistas"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Padrón de Proveedores"
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
            TabIndex        =   61
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   62
         Top             =   525
         Width           =   1905
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
            Picture         =   "OFCONF04.frx":3CFE
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Maestro de Artículos"
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
            Top             =   105
            Width           =   1065
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
         TabIndex        =   12
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Programación de Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5040
      TabIndex        =   10
      Top             =   3960
      Width           =   4425
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   52
         Top             =   525
         Width           =   1800
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
            TabIndex        =   53
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   50
         Top             =   525
         Width           =   1800
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
            TabIndex        =   51
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base de Datos de O-F's  y  Envios"
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
      Left            =   180
      TabIndex        =   2
      Top             =   1560
      Width           =   4955
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   920
         ScaleHeight     =   585
         ScaleWidth      =   3270
         TabIndex        =   19
         Top             =   420
         Width           =   3330
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
            Left            =   3840
            Picture         =   "OFCONF04.frx":4008
            Style           =   1  'Graphical
            TabIndex        =   55
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
            Left            =   2740
            Picture         =   "OFCONF04.frx":4312
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Recepción de Envio"
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
            Left            =   2200
            Picture         =   "OFCONF04.frx":445C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anulación de Envio"
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
            Left            =   1680
            Picture         =   "OFCONF04.frx":4766
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Nuevo Envio de Prendas"
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
            Picture         =   "OFCONF04.frx":48B0
            Style           =   1  'Graphical
            TabIndex        =   22
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
            Left            =   530
            Picture         =   "OFCONF04.frx":4BBA
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cierre de Orden de Fabricación - Re-Apertura O/F's con Saldo"
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
            Left            =   1060
            Picture         =   "OFCONF04.frx":4EC4
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Anulación de Orden de Fabricación"
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
            Picture         =   "OFCONF04.frx":51CE
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Lanzamiento de Orden de Fabricación"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2720
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   4425
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   92
         Top             =   1200
         Width           =   3795
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
            Left            =   3150
            Picture         =   "OFCONF04.frx":54D8
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "4"
            Top             =   0
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
            Left            =   2550
            Picture         =   "OFCONF04.frx":591A
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "3"
            Top             =   0
            Width           =   600
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
            Left            =   1950
            Picture         =   "OFCONF04.frx":5A64
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "2"
            Top             =   0
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
            Left            =   1365
            Picture         =   "OFCONF04.frx":5D6E
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "1"
            Top             =   0
            Width           =   600
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
            Height          =   600
            Left            =   0
            Picture         =   "OFCONF04.frx":5EB8
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Consulta Base de Datos Completa"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "B-D de Envios "
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
            TabIndex        =   93
            Top             =   105
            Width           =   750
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
         Left            =   315
         TabIndex        =   47
         Top             =   1920
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
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   41
         Top             =   420
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
            Picture         =   "OFCONF04.frx":61C2
            Style           =   1  'Graphical
            TabIndex        =   56
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
            Picture         =   "OFCONF04.frx":64CC
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Consulta por Número de Orden de Fabricación"
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
            Picture         =   "OFCONF04.frx":67D6
            Style           =   1  'Graphical
            TabIndex        =   44
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
            Picture         =   "OFCONF04.frx":6AE0
            Style           =   1  'Graphical
            TabIndex        =   43
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
            Picture         =   "OFCONF04.frx":6C2A
            Style           =   1  'Graphical
            TabIndex        =   42
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
            TabIndex        =   45
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Complementos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   4955
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   39
         Top             =   525
         Width           =   2000
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
            Left            =   735
            TabIndex        =   40
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2600
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   36
         Top             =   525
         Width           =   2000
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
            Picture         =   "OFCONF04.frx":6F34
            Style           =   1  'Graphical
            TabIndex        =   37
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
            Left            =   840
            TabIndex        =   38
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
      X2              =   10395
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Orfab 
      Caption         =   "Ordenes de Fabricacion"
      Begin VB.Menu NvaOrd 
         Caption         =   "Nueva Orden"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LanOrd 
         Caption         =   "Lanzar Orden"
      End
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar Orden"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu SusOrd 
         Caption         =   "Suspensión de Orden"
      End
      Begin VB.Menu AnuOrd 
         Caption         =   "Anulación de Orden"
      End
      Begin VB.Menu ModOrd 
         Caption         =   "Modificación de Orden"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu VerImp 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Begin VB.Menu SuOrd 
         Caption         =   "Suspender Orden"
      End
      Begin VB.Menu AnulOrd 
         Caption         =   "Anular Orden"
      End
      Begin VB.Menu ModiOrd 
         Caption         =   "Modificar Orden"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimpi 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu AsiMaPri 
      Caption         =   "Materiales"
      Begin VB.Menu MatReser 
         Caption         =   "Reserva de Materiales"
         Begin VB.Menu GenReservOC 
            Caption         =   "Generar Reservas por Orden de Fabricación"
         End
         Begin VB.Menu CanceReser 
            Caption         =   "Cancelación de Reservas"
         End
         Begin VB.Menu ConOrFaCu 
            Caption         =   "Control Segun Ordenes de Fabricación en Curso"
         End
         Begin VB.Menu DepuBdD 
            Caption         =   "Depuración de Base de Datos"
         End
         Begin VB.Menu php3a 
            Caption         =   "-"
         End
         Begin VB.Menu LiToMat 
            Caption         =   "Listado de Totales por Material"
            Begin VB.Menu LiToPan 
               Caption         =   "por Pantalla"
            End
            Begin VB.Menu LiToImpre 
               Caption         =   "por Impresión"
            End
         End
         Begin VB.Menu CoGrlRe 
            Caption         =   "Consulta General de Reservas por Orden de Fabricación"
         End
         Begin VB.Menu ReXNumOF 
            Caption         =   "Reservas por Numero de Orden de Fabricación"
         End
         Begin VB.Menu ReXMat 
            Caption         =   "Reservas por Material"
         End
      End
      Begin VB.Menu ValMate 
         Caption         =   "Vale de Materiales"
         Begin VB.Menu ValPrep 
            Caption         =   "Vale de Preparacion"
         End
         Begin VB.Menu ValTran 
            Caption         =   "Vale de Transferencia"
         End
      End
   End
   Begin VB.Menu AcDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu InfoPro 
         Caption         =   "Informes de Producción"
      End
      Begin VB.Menu CarMaq 
         Caption         =   "Carga de Maquinas"
      End
      Begin VB.Menu php15 
         Caption         =   "-"
      End
      Begin VB.Menu MaestArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu php16 
         Caption         =   "-"
      End
      Begin VB.Menu otraccdir 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu ConsuLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ConGrBdD 
         Caption         =   "Consulta General de la Base de Datos"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu PenSecPro 
         Caption         =   "Pendientes por Sector Productivo"
         Begin VB.Menu Xpanta 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpres 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu PenNumOF 
         Caption         =   "Pendientes por Numero de Orden de Fabricación"
      End
      Begin VB.Menu PenCodPro 
         Caption         =   "Pendientes por Código de Producto"
         Begin VB.Menu porPantall 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpreso 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConSelEst 
         Caption         =   "Consulta Selectiva por Estado"
         Begin VB.Menu Xpant 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpre 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimp 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu php24 
         Caption         =   "-"
      End
      Begin VB.Menu ProCar 
         Caption         =   "Programación de Carga"
         Begin VB.Menu OperPro 
            Caption         =   "Operaciones Programadas"
         End
         Begin VB.Menu CaToEqui 
            Caption         =   "Carga Total por Equipo"
         End
      End
   End
   Begin VB.Menu ConfGral 
      Caption         =   "Configuración"
   End
End
Attribute VB_Name = "OFCONF04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AnulOrd_Click()
Call Command45_Click
End Sub

Private Sub AnuOrd_Click()
Call Command45_Click
End Sub

Private Sub CanceReser_Click()
Call RESERV04.Command4_Click
End Sub

Private Sub CarMaq_Click()
Call Command52_Click
End Sub

Private Sub CaToEqui_Click()
Call Command6_Click
End Sub

Private Sub Cerrar_Click()
Call Command46_Click
End Sub

Private Sub CoGrlRe_Click()
'
RESERV04.Option1 = True
RESERV04.Option5 = True
Call RESERV04.Command2_Click
'

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call GENORDEFDAT(1)
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C03", "Consulta O/F's por Sector")
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
   Call VALEMATER(0)
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   HOII% = HOY(HOS$)
10 VDEF$ = MKI$(COMESACT%) + MKI$(HOII%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   '
   FEX = DES%: DESX$ = FECHATEX$(FEX)
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   '
   JJ& = 0
   Screen.MousePointer = 11
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MOVISTO", U&)
     FF% = CVI(Left$(XX$, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         If CVD(Mid$(XX$, 88, 8)) > 0.001 Then
           CMOX$ = TRIM$(UCase$(Mid$(XX$, 21, 2)))
           If CMOX$ = "OF" Or CMOX$ = "PF" Then
             YY$ = REGBLAN$("DETAFABR")
             Call REPLA(YY$, MKI$(FF%), 1, 2)
               CIXI% = CVI(Mid$(XX$, 3, 2))
             Call REPLA(YY$, MKI$(CIXI%), 3, 2)
               NDOCUM$ = TRIM$(Mid$(XX$, 23, 10))
             Call REPLA(YY$, NDOCUM$, 5, 10)
               CAENT# = CVD(Mid$(XX$, 88, 8))
             Call REPLA(YY$, MKD$(CAENT#), 15, 8)
               REFEXXX$ = TRIM$(Mid$(XX$, 43, 30))
             Call REPLA(YY$, REFEXXX$, 23, 30)
             '
             JJ& = JJ& + 1
             Call GRAREG("DETAFABR", YY$, JJ&)
             '
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("DETAFABR", JJ&)
   Call CIERRARCH("DETAFABR")
   '
   Call HEADERS("DETAFABR", "")
   Call HEADERS("DETAFABR", "Periodo: " + DESX$ + " - " + HASX$)
   '
   Call FILESORT("DETAFABR", "", 1, 2, "ASC")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Call FINAL("*DETAFABR")
   '
99 Command12.Enabled = True
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   OPTEROPER.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
'    OPORFA04.Show 1
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call VERORFA
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Padrón de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If EXISTE%("AMASTO02.EXE") > 0 Then
         Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
       Else
         Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
    End If
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   ACONEC$ = "REPARA02"
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "ARCON") > 0 Then ACONEC$ = "REPABO03"
   Call CONECRUN(ACONEC$, "Reparaciones")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call REPRINORFA
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
    Call ENVPREN 'NUEVO ENVIO DE PRENDA
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
    Command41.Enabled = False
    Call ORFNUE(1)
    Command41.Enabled = True
End Sub

Private Sub Command49_Click()
   Command49.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Command49.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call VERORFA
   Command5.Enabled = True
End Sub

Private Sub Command51_Click()
    Command51.Enabled = False
    ACONEC$ = "IFABRI04"
    Call CONECRUN(ACONEC$, "Informes de Producción")
    Command51.Enabled = True
End Sub

Private Sub Command52_Click()
    Command52.Enabled = False
    ACONEC$ = "CARMAQ04"
    Call CONECRUN(ACONEC$, "Carga de Máquinas")
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
    Call GENORDEFDAT(1)
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
               CIXI% = CVI(Mid$(X$, 83, 2))
               Call REPLA(X$, MKI$(CIXI%), 1, 2)
               Call REPLA(X$, MKI$(0), 3, 2)
               Call REPLA(X$, Space$(64), 5, 64)
               CATOT# = CVD(Mid$(X$, 89, 8))
               CAPRA# = CVD(Mid$(X$, 119, 8))
               CAPEN# = CATOT# - CAPRA#
               If CAPEN# < 0 Then CAPEN# = 0
               Call REPLA(X$, MKD$(CAPEN#), 159, 8)
               Call REPLA(X$, MKD$(0), 167, 8)
               J& = J& + 1
               Call GRAREG("ORFAMAT", X$, J&)
             End If
           End If
         Next U&
         Call SETULTREG("ORFAMAT", J&)
         Call FILESORT("ORFAMAT", "", 1, 4, "ASC")
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         '
         Call FINAL("*LORFAMAT")
    End If
    Command64.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call ACTECOR(1)
    Call GENEOPERFAP
    Call FINAL("*OPERFAP")
    Screen.MousePointer = 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    RESERV04.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Call ORFANU
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    OPX% = COMALTER%("Opciones de Trabajo con O/F's:\\Cierre de O/F's\Re-Apertura")
    If OPX% = 1 Then
        Call ORFCIE
      ElseIf OPX% = 2 Then
        If DERACCE%("REAPEROF", "Re-Apertura O/F's con Saldo") >= 2 Then
          Call REAPEROF
        End If
    End If
    Command46.Enabled = True
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    Call REVIORFLAN
    If Option1.Value = True Then
           Call GENORDEFDAT(0)
           Call CONECRUN("*CONORFA", "Consulta Base de datos de O.F's")
        Else
           Call GENORDEFDAT(0)
           Call OPNOIN("")
    End If
    Command61.Enabled = True
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
    Call GENORDEFDAT(0) 'NUEVO
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
            DESI% = (-32767): HASI% = 32767: PERIO$ = ""
            If Asc(OPX$) = 3 Then
              HOII% = HOY(HOS$)
10            VDEF$ = MKI$(COMESACT%) + MKI$(HOII%)
              OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
              If OBX$ = "" Then GoTo 99
              '
              DESI% = CVI(Left$(OBX$, 2))
              HASI% = CVI(Mid$(OBX$, 3, 2))
              If HASI% < DESI% Then GoTo 10
              '
              FEX = DESI%: DESX$ = FECHATEX$(FEX)
              FEX = HASI%: HASX$ = FECHATEX$(FEX)
              PERIO$ = " - Desde: " + DESX$ + " Hasta: " + HASX$
            End If
            '
            J& = 0
            Screen.MousePointer = 11
            Call HEADERS("ORFASEL", "")
            Call HEADERS("ORFASEL", "Estado: " + ECOARCH$("ORFASTA", OPX$) + PERIO$)
            For U& = 1 To ULTREG&("ORDEFABR")
              X$ = REGLEIDO$("ORDEFABR", U&)
              If Mid$(X$, 108, 1) = OPX$ Then
                FECIX% = CVI(Mid$(X$, 117, 2))
                If FECIX% >= DESI% Then
                  If FECIX% <= HASI% Then
                    J& = J& + 1
                    Call GRAREG("ORFASEL", X$, J&)
                  End If
                End If
              End If
            Next U&
            Call SETULTREG("ORFASEL", J&)
            Call CIERRARCH("*.*")
            Screen.MousePointer = 1
            '
            If J& < 1 Then
                Call MENSERR(24, "Sin Registros para Criterio de Selección")
              Else
                Call FINAL("*LIORFST")
            End If
            '
          End If
        End If
    End If
99  Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    If Option1.Value = True Then
           '
           Call GENORDEFDAT(0)
           Screen.MousePointer = 11
           Call ABREORFAB
           SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
           SQLX$ = SQLX$ + "where CanApro < 0.05 "
           SQLX$ = SQLX$ + "AND CanRech < 0.05 "
           SQLX$ = SQLX$ + "order by IdSubOr ASC"
           Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
           '
           On Error Resume Next
           With ORFATEMP
             .MoveFirst
             If Err Then
               Call MENSERR(24, "No Hay O/F's Pendientes\que Puedan Refrescarse.")
               Screen.MousePointer = 1
               Exit Sub
             End If
             On Error GoTo 0
             '
             JJ& = 0
             Do While Not .EOF
               If Val(Mid$(!IdSubOr, 11)) = 1 Then
                 ZZ$ = REGBLAN$("ORFANUE")
                 REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
                 Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
                 Call REPLA(ZZ$, REFXX$, 5, 64)
                 JJ& = JJ& + 1
                 Call GRAREG("ORFANUE", ZZ$, JJ&)
               End If
               .MoveNext
             Loop
           End With
           Call SETULTREG("ORFANUE", JJ&)
           Call CIERRARCH("ORFANUE")
           Call FRESHECHO("ORFANUE")
           Screen.MousePointer = 1
           '
           Call CONECRUN("*ORF-C01", "Consulta por Número de O.F.")
        Else
           Call GENORDEFDAT(0)
           Call OPNOIN("")
    End If
    Command63.Enabled = True
End Sub

Private Sub ConfGral_Click()
Call Command14_Click
End Sub

Private Sub ConGrBdD_Click()
Option1.Value = True
Call Command61_Click
End Sub

Private Sub ConOrFaCu_Click()
Call RESERV04.Command7_Click
End Sub

Private Sub DepuBdD_Click()
Call RESERV04.Command5_Click
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
    'Show
    '
    VERANTER$ = "ORDEFABR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If APLINVO$ = "OPRESERV" Then
      RESERV04.Show 1
      Call FINAL("")
    End If
    '
    If APLINVO$ = "PRODUPRO" Then
      Option2.Value = True
      Call Command64_Click
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
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
      If EXISTE%("REPABO02.EXE") < 1 Then
        Command3.Enabled = False
        Label5.Enabled = False
      End If
    End If
    '
    Show
    '
End Sub
'

Sub VALEMATER(MODOVA%)
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   If MODOVA% < 1 Or MODOVA% > 2 Then
      MODOVA% = ALTERNA%("Vale de Preparacion\Vale de Transferencia")
   End If
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
   SecProd$ = Mid$(XX$, 5, 32)
   DEPODEST% = Asc(Mid$(XX$, 81, 1))
   '
   DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
   If DEPORIG% < 1 Then
      Call MENSERR(24, "Falta Definir Depósito Central")
      GoTo 13
   End If
   '
   EPAC$ = EMPREAC$
   If InStr(EPAC$, "DOSIVAC") > 0 Then
13    Call SELECHO("TADEPOSI")
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
   Call GENORDEFDAT(1)
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
   Call RESERV04.GENRESDAT
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
   FIN& = NUMAS%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      CI% = U&
      II1$ = CODEXT$(CI%)
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
   Next U&
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1
   '
   If MODOVA% = 1 Then
      TITUPA$ = "Vale de PREPARACION" + " - de " + DEPOSI$ + " a " + SecProd$
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
      Call HEADERS("VALEMAT", "Seccion: " + SECCX$ + " - " + SecProd$)
      Screen.MousePointer = 1
      Call CONECRUN("*PRIVALMA", "Vale de Preparacion de Materiales")
      GoTo 10
   End If
   '
   TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + DEPOSI$ + " a " + SecProd$
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
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + SecProd$
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPORIG%, CANTI)
         '
         REFE$ = "TR.de " + DEPOSI$
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
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
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = CSTRING$(MKS$(SECCI), 3, 8, 3, 2) + " - " + SecProd$
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
    'NUEVO
      Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, "NO Existen Ordenes de Fabricación. Imposible Continuar la Carga.")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEF")
        '
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEF", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("ECORDEF", JJ&)
    Call CIERRARCH("ECORDEF")
    'HASTA ACA NUEVO
    '
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
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
    EDITFORM.Command1.Enabled = False
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

Sub ORFNUE(REPIM%)                           ' NUEVA ORDEN DE FABRICACION
    '
    ' REPIM%=1 ==> nueva orden de fabricacion
    ' REPIM%=2 ==> repite impresion de O/FAB
    '
    Static YADEPU%
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    Dim CILOTE%(256), CALOTE#(256), CALOTE1#(256), PUNMAT$(256), PUNOPE$(256)
    Dim CISUMI%(512), CASUMI#(512)
    '
    CANCELPRINT% = 0
    If CONTROL$("ORDEFABR", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
    GoSub FORIPRE                      ' formulario de impresion
    GoSub MONUME                       ' modo de numeracion
    GoSub GESUBORDEN                   ' genera sub-ordenes
    GoSub OBSERVACI                    ' incluye observaciones
    '
100 If REPIM% = 2 Then
       Screen.MousePointer = 11
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
       SQLX$ = SQLX$ + "WHERE StatuOrf <= 1 "
       SQLX$ = SQLX$ + "AND CanApro < 0.05 "
       SQLX$ = SQLX$ + "AND CanRech < 0.05 "
       SQLX$ = SQLX$ + "order by IdSubOr ASC"
       Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       On Error Resume Next
       With ORFATEMP
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No Hay O/F's Pendientes\que Puedan Refrescarse.")
           Screen.MousePointer = 1
           Exit Sub
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While Not .EOF
           If Val(Mid$(!IdSubOr, 11)) = 1 Then
             ZZ$ = REGBLAN$("ORFANUE")
             REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
             Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
             Call REPLA(ZZ$, REFXX$, 5, 64)
             JJ& = JJ& + 1
             Call GRAREG("!ORFANUE", ZZ$, JJ&)
           End If
         .MoveNext
         Loop
       End With
       Call SETULTREG("!ORFANUE", JJ&)
       Call CIERRARCH("!ORFANUE")
       Call FRESHECHO("!ORFANUE")
       Screen.MousePointer = 1
       '
       Call SELECHO("!ORFANUE/Ordenes de Trabajo para Re-Generación")
       NRO& = Val(VALACT1$("!ORFANUE"))
       If NRO& < 1 Then
          INDI% = 0
          Call CIERRARCH("*.*")
          Call LIBARCH("*.*")
          Exit Sub
       End If
    End If
    '
    If MODORFA% = 0 Then
       GoSub MODOGEN                      ' modo directo / MRP
    End If
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("RESERVA")
    Call FRESHECHO("ORFAEMI")
    '
    If REPIM% = 2 Then
       Call ABREORFAB
       SQLX$ = "SELECT * FROM ORDEFABR "
       SQLX$ = SQLX$ + "WHERE NumeOrFa = " & NRO& & " "
       SQLX$ = SQLX$ + "order by IdSubOr ASC"
       Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       On Error Resume Next
       With ORFATEMP
         .MoveFirst
         If Err Then GoTo 1005
         On Error GoTo 0
         '
         JJ& = 0: CIOF% = 0: CAN# = 0
         Do While Not .EOF
            ' X1$ = X3$: X2$ = X3$: TORFANUE$ = X3$
            If CIOF% = 0 Then
              CIOF% = !Cod_Inte
              LOTE# = !CanProy
              Status% = !Statuorf
            End If
            If !Cod_Inte = CIOF% Then CAN# = CAN# + !CanProy
            FGEN = CVINT(!FechGen)
            FLAN% = CVINT(!FechLan)
            FENTRE% = CVINT(!FechEnSol)
              If FENTRE% < HOY(HOS$) Then FENTRE% = HOY(HOS$)
            SPX$ = MKS$(!SecciPro)
            DESTI$ = !Destino
            If !Statuorf > 1 Then
              Call MENSERR(24, "Imposible Regenerar.\O/F Cumplida, Suspendida o Anulada.")
              GoTo 100
            End If
         .MoveNext
         Loop
       End With
       CI% = CIOF%: GoTo 104
       '
1005   On Error GoTo 0
       Call MENSERR(24, "Orden no Encontrada")
       GoTo 100
    End If
    '
    GoSub ARMAPRIMERA
    X1$ = OBJPLA$(OBJEPRI$, X0$)
    If X1$ = "" Then GoTo 999
    '
    GoSub VERICODI
    If CIOF% < 1 Then GoTo 100
    CIXI% = CIOF%
    SPX$ = String$(4, 0): SPY$ = ""
    '
    If ULTREG&("SECPROD") > 0 Then
101   OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + MKS$(SEPRODUC(CIXI%)))
      If OBX$ = "" Then GoTo 999
      SPX$ = Mid$(OBX$, 3, 4)
      If EXISEPRO% = 1 Then
        If ECOARCH$("SECPROD", SPX$) = "" Then GoTo 101
      End If
      SPY$ = ""
      If CVS(SPX$) > 0 Then
        SPY$ = TRIM$(CSTRING$(SPX$, 3, 8, 3, 2))
      End If
    End If
    '
    GoSub AJUCANORF
' HASTA AQUI LLEGUE ***************************************************
    GoSub ARMASEGUNDA
    TMT% = TIMATE%(CIOF%)   ' PARA PONER TIPO DE MATERIAL EN POSICION 257,1
102 X2$ = Left$(OBJPLA$("ORFANUE-2", X1$ + Chr$(TMT%)), 256)
    If X2$ = "" Then GoTo 100
    '
    GoSub VALIDA1                   ' validacion de entrada
    If VALIDA% < 0 Then
         GoTo 100
       ElseIf VALIDA% = 0 Then
         GoTo 102
    End If
    '
    If VERITAND% = 1 Then
       If LOECON >= 1 Then
          If CANORF / LOECON <> Int(CANORF / LOECON) Then
             Call MENSERR(24, "La Cantidad Ordenada no Resulta\un Múltiplo Entero de Lote Standard (Tanda).\  \Ajustada a un Múltiplo de Tanda.")
             X1$ = X2$
             CANORF = LOECON * Int(CANORF / LOECON + 0.5)
             Call REPLA(X1$, MKD$(CDbl(CANORF)), 89, 8)
             GoTo 102
          End If
       End If
    End If
    '
103 XACCE$ = ""
    GoSub OBSERVOF                  ' ingresa observaciones
    If VALIDA% < 0 Then
       GoTo 100
    End If
    '
    'Call CIERRARCH("ORFANUE")
    Call CIERRARCH("*.*")
    '
104 GoSub GELOTES       ' genera subordenes por lotes
    GoSub GESUBOR       ' genera subordenes por subconjuntos
    '
    For II% = 1 To NORDE%
      If CANOPS%(CILOTE%(II%)) < 1 Then
        CMAT% = COMALTER%("Falta Definir Operaciones para\" + TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT0$(CILOTE%(II%))) + "\\Cancelar\Continuar")
        If CMAT% <> 2 Then GoTo 102
      End If
    Next II%
    '
    GoSub GERESMAT      ' genera reserva de componentes
       If VACONTROL% < 0 Then
         If REPIM% = 2 Then
            GoTo 100
          Else
            GoTo 102
         End If
       End If
       '
    'Call CIERRARCH("ORFANUE")
    Call CIERRARCH("*.*")
    '
    'Call BLOQARCH("ORDEFABR")
    'Call BLOQARCH("INDIORF")
    'Call BLOQARCH("RESERVA")
    'Call BLOQARCH("OPERFAB")
    'Call BLOQARCH("OPERFAP")
    'Call BLOQARCH("ECORDEP")
    'Call BLOQARCH("ECORDEF")
    'Call BLOQARCH("ORFAEMI")
    'Call BLOQARCH("ORFANUE")
    '
    TERENUME$ = ""
    If REPIM% <> 2 Then
       '
       NROPROX& = 1
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
       SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
       Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       On Error Resume Next
       With ORFABTEMP
         .MoveFirst
         If Err < 1 Then
           NROPROX& = 1 + !NUMEORFA
         End If
       End With
       On Error GoTo 0
       '
       If NROPROX& <> NRO& Then
          TERENUME$ = "FLEXOFT ha Re-Numerado la O/F\como Orden Número '" + TRIM$(Str$(NROPROX&)) + "'."
          NRO& = NROPROX&
       End If
    End If
    '
    GoSub GRAORFNUE
    Call CIERRARCH("ORFANUE")
    GoSub GRAORDEFA
    Call CIERRARCH("ORDEFABR")
    '
    GoSub GRARESERV
    GoSub DEPUORFEM
    '
    GoSub GRASECOP
    '
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("ORDEFABR")
    Call CIERRARCH("INDIORF")
    Call CIERRARCH("RESERVA")
    Call CIERRARCH("OPERFAB")
    Call CIERRARCH("OPERFAP")
    Call CIERRARCH("ECORDEP")
    Call CIERRARCH("ECORDEF")
    Call CIERRARCH("ORFAEMI")
    '
    GoSub PRINORFAB
    '
    If TERENUME$ <> "" Then
       Call COMUNI(TERENUME$)
    End If
    '
    GoTo 100
    '
    '
' ************************************* subrutinas asociadas
    '
FORIPRE:
    FORIPRE$ = TRIM$(CONTROL$("", "FORMORFA"))
    If FORIPRE$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Fabricación:\  \'" + RMCC$ + "'.")
            FORIPRE$ = ""
        End If
    End If
    '
    FORIPRE2$ = TRIM$(CONTROL$("", "FORMORFB"))
    If FORIPRE2$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE2$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE2$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Fabricación:\  \'" + RMCC$ + "'.")
            FORIPRE2$ = ""
        End If
    End If
    Return
    '
MONUME:
    MONUME$ = "AUTO"
    If CONTROL$("", "NUMEORFA") = "MANUAL" Then
        MONUME$ = "MANUAL"
    End If
    Return
    '
GESUBORDEN:
    GESUBOR% = 0
    If CONTROL$("", "GESUBOR") = "SI" Then
        GESUBOR% = 1
    End If
    '
    VERITAND% = 0
    If CONTROL$("", "VERITAND") = "SI" Then
        VERITAND% = 1
    End If
    '
    EXISEPRO% = 0
    If CONTROL$("", "EXISEPRO") = "SI" Then
        EXISEPRO% = 1
    End If
    '
    Return
    '
OBSERVACI:
    OBSERVOF% = 0
    If CONTROL$("", "OBSERVOF") = "SI" Then
        OBSERVOF% = 1
    End If
    Return
    '
MODOGEN:
    If PRECOLAN% = 0 Then
       PRECOLAN% = (-1)
       If CONTROL$("", "PRECOLAN") = "SI" Then
          PRECOLAN% = 1
       End If
    End If
    '
    If INCLUPRI% = 0 Then
       INCLUPRI% = (-1)
       If CONTROL$("ORFABRI", "MPRIUNIQ") = "INCLUPRI" Then
          INCLUPRI% = 1
       End If
    End If
    '
    If FRALOMAX% = 0 Then
       FRALOMAX% = (-1)
       If CONTROL$("", "FRALOMAX") = "SI" Then
          FRALOMAX% = 1
       End If
    End If
    '
    If VALINDOP% = 0 Then
       VALINDOP% = (-1)
       If CONTROL$("", "VALINDOP") = "SI" Then
          VALINDOP% = 1
       End If
    End If
    '
    MODORFA% = 1
    If REPIM% <> 2 Then
401    MODORFA% = COMALTER%("Puede Generar Ordenes de Fabricación\en Forma Directa, sobre la Base\del Cálculo MRP o de una O/Venta\\Directa\MRP\O/Venta\Cancelar")
       If MODORFA% = 1 Then
             OBJEPRI$ = "ORFANUE-1"
          ElseIf MODORFA% = 2 Then
             OBJEPRI$ = "ORFANUE-9"
          ElseIf MODORFA% = 3 Then
             Call MENSERR(24, "No Disponible en la Presente Versión")
             GoTo 401
             '
             'Call GENPEDPEN
             'Call SELECHO("INDIPEP/Indice General de Pedidos Pendientes")
             'NPX& = Val(VALACT1$("INDIPEP"))
             'If NPX& <> 0 Then
             '  Ofabric.Label1.Caption = CStr(NPX&)
             '  FORSELPPE.Show 1
             '  RFFS$ = Ofabric.Label1
             '  RFFA$ = Ofabric.Label12
             '  Call CIERRARCH("*.*")
             '  Exit Sub
             'End If
          Else
             Call CIERRARCH("*.*")
             Exit Sub
       End If
       If MONUME$ = "MANUAL" Then
          OBJEPRI$ = OBJEPRI$ + "M"
       End If
    End If
    Return
    '
'**************************************************************************
ARMAPRIMERA:
    '
    NRO& = 1
    Call ABREORFAB
    '
    SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
    SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
    Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With ORFABTEMP
      .MoveFirst
      If Err < 1 Then
        NRO& = 1 + !NUMEORFA
      End If
    End With
    On Error GoTo 0
    '
    X0$ = REGBLAN$("ORFANUE")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 69, 2)
    Call REPLA(X0$, MKI$(HO%), 109, 2)
    Call REPLA(X0$, MKI$(HO%), 113, 2)
    Return
    '
'**************************************************************************
VERICODI:                                                         ' esta OK
    CIOF% = 0: UMED$ = "U.": LOSTA = 1
    CI% = CVI(Mid$(X1$, 71, 2))
    If CI% < 1 Then
         Call MENSERR(24, "Falta Código de Producto.")
       Else
         CIOF% = CI%
         UMED$ = UNIMED$(CI%)
         LOSTA = LOREPOMA(CI%)
         If LOSTA < 1 Then LOSTA = LOTESTAN(CI%)
    End If
    Return
    '
'**************************************************************************
BORENACCE:
    XACCE$ = ""
    RGACC& = 0
    For U& = 1 To ULTREG&("BOREACCE")
       YCC$ = REGLEIDO$("BOREACCE", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          RGACC& = U&
          YACCE$ = YCC$
          GoTo 710
       End If
    Next U&
    '
    YACCE$ = String$(128, 0)
    Call REPLA(YACCE$, MKI$(CIOF%), 1, 2)
    Call REPLA(YACCE$, Space$(28), 3, 28)
    Call REPLA(YACCE$, Space$(40), 63, 40)
    Call REPLA(YACCE$, Space$(18), 105, 16)
    '
710 Call REPLA(YACCE$, TALLE$(CIOF%), 3, 12)
    Call REPLA(YACCE$, COLOR$(CIOF%), 15, 16)
    '
    RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIOF%))
    If RFF$ <> "" Then
       REGEST& = CVS(Left$(RFF$, 4))
       CIHIL% = CVI(Mid$(REGLEIDO$("ESTRUNUE", REGEST&), 3, 2))
       Call REPLA(YACCE$, TIPOHIL$(CIHIL%), 63, 12)
       Call REPLA(YACCE$, TITULO$(CIHIL%), 75, 12)
       Call REPLA(YACCE$, COLOR$(CIHIL%), 87, 16)
    End If
    '
    XACCE$ = YACCE$
    If REPIM% <> 2 Then
       XACCE$ = OBJPLA$("BORENACCE", YACCE$)
    End If
    '
    If XACCE$ <> "" Then
       PROVHIL$ = TEL2CLI$((-1) * CVI(Mid$(XACCE$, 103, 2)))
       PARTHIL$ = Mid$(XACCE$, 105, 16)
       '
       LARCUTE$ = CSTRING$(Mid$(XACCE$, 31, 4), 3, 8, 1, 1)
       ACHCUTE$ = CSTRING$(Mid$(XACCE$, 35, 4), 3, 8, 1, 1)
       LARMATE$ = CSTRING$(Mid$(XACCE$, 39, 4), 3, 8, 1, 1)
       ACHMATE$ = CSTRING$(Mid$(XACCE$, 43, 4), 3, 8, 1, 1)
       LARCUCO$ = CSTRING$(Mid$(XACCE$, 47, 4), 3, 8, 1, 1)
       ACHCUCO$ = CSTRING$(Mid$(XACCE$, 51, 4), 3, 8, 1, 1)
       LARMACO$ = CSTRING$(Mid$(XACCE$, 55, 4), 3, 8, 1, 1)
       ACHMACO$ = CSTRING$(Mid$(XACCE$, 59, 4), 3, 8, 1, 1)
       '
       If RGACC& < 1 Then RGACC& = 1 + ULTREG&("BOREACCE")
       Call GRAREG("BOREACCE", XACCE$, RGACC&)
    End If
    Call CIERRARCH("BOREACCE")
    '
    Return

BORENCOMPLE:
    Call BLANARCH$("!BCOMPLEM")
    HAYCOMP% = 0
    For U& = 1 To ULTREG&("BCOMPLEM")
       YCC$ = REGLEIDO$("BCOMPLEM", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          CACOM = CVS(Mid$(YCC$, 49, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 53, 4)
          CACOM = CVS(Mid$(YCC$, 57, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 61, 4)
          CACOM = CVS(Mid$(YCC$, 69, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 73, 4)
          Call REGAPP("!BCOMPLEM", YCC$)
          HAYCOMP% = 1
       End If
    Next U&
    '
    If HAYCOMP% = 0 Then
       YCC$ = REGBLAN$("BCOMPLEM")
       Call REPLA(YCC$, MKI$(CIOF%), 1, 2)
       Call REPLA(YCC$, "Cuerpo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Manga", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Puño", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Cuello", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Tiras", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Capucha", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Bolsillo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Presilla", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
    End If
    '
    REFCOMPX$ = ""
    TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODEXT$(CIOF%)) + " - " + DESCRIT0$(CIOF%)
    '
    VTB% = 0
    If REPIM% <> 2 Then
       VTB% = VENTABU%("BCOMPLEM/TITUPAN=Orden de Complementos:" + TITU$)
    End If
    '
    If VTB% >= 0 Then
       If HAYCOMP% > 0 Then
           J& = 0
           For U& = 1 To ULTREG&("BCOMPLEM")
             YCC$ = REGLEIDO$("BCOMPLEM", U&)
             If CVI(Left$(YCC$, 2)) <> CIOF% Then
                J& = J& + 1
                Call GRAREG("BCOMPLEM", YCC$, J&)
             End If
           Next U&
         Else
           J& = ULTREG&("BCOMPLEM")
       End If
       For U& = 1 To ULTREG&("!BCOMPLEM")
          YCC$ = REGLEIDO$("!BCOMPLEM", U&)
          CACOM = CVS(Mid$(YCC$, 53, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOM), 49, 4)
          CACOX = CVS(Mid$(YCC$, 61, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 57, 4)
          CACOX = CVS(Mid$(YCC$, 73, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 69, 4)
          Call GRAREG("!BCOMPLEM", YCC$, U&)
          J& = J& + 1
          Call GRAREG("BCOMPLEM", YCC$, J&)
          '
          If CACOM > 0 Then
             If REFCOMPX$ <> "" Then REFCOMPX$ = REFCOMPX$ + " - "
             REFCOMPX$ = REFCOMPX$ + TRIM$(Mid$(YCC$, 3, 14))
          End If
       Next U&
       Call SETULTREG("BCOMPLEM", J&)
    End If
    Return
    '
'**************************************************************************
AJUCANORF:                                                        ' esta OK
    CAMRP = 0
    CANORF = LOSTA
    LOECON = LOTESTAN(CIOF%)
    If MODORFA% = 2 Then
        CAMRP = CVD(Mid$(REGACT$("ORFAEMI"), 49, 8))
        CANORF = CAMRP
        If CAMRP < LOECON Then
          CANORF = LOECON
          Call MENSERR(24, "La Cantidad Pedida es Inferior\a un Lote Económico.\  \Ajustada a un Lote Standard Económico.")
        End If
      ElseIf MODORFA% = 1 Then
        CANORF = LOECON
        If CANORF < 1 Then CANORF = 1
    End If
    '
    If LOSTA < CANORF Then
      XLMX$ = MKI$(CI%) + MKS$(CANORF)
      XLMY$ = OBJPLA$("CARLOREPOMA", XLMX$)
      If XLMY$ <> "" Then
        LOSTA = CVS(Mid$(XLMY$, 3, 4))
        If LOSTA < 1 Then LOSTA = 1
        Master.FindFirst ("CODINT = " & CIOF%)
        If Master.NoMatch = False Then
          Master.Edit
          Master!LOREPOS = LOSTA
          Master.Update
        End If
      End If
    End If
    Return
    '
'**************************************************************************
ARMASEGUNDA:
    NRO& = CVS(Left$(X1$, 4))
    REOFE& = 0
    FENTRESOL% = HO%
    If MODORFA% = 2 Then
        FENTRESOL% = CVI(Mid$(REGACT$("ORFAEMI"), 57, 2))
        REOFE& = REGNUM&("ORFAEMI")
    End If
    '
    FENTREGA% = HO%
    If MODORFA% = 2 Then FENTREGA% = FENTRESOL%
    If FENTREGA% < HO% Then
        FEX = FENTREGA%
        Call MENSERR(24, "Imposible Cumplir la Fecha de\Entrega Solicitada del " + FECHATEX$(FEX) + ".\  \Entrega Re-Programada para el " + HOS$ + ".")
        FENTREGA% = HO%
    End If
    '
    DESTI$ = "STOCK": If TIMATE%(CIOF%) = 4 Then DESTI$ = "SEMIELA"
    'If SPY$ <> "" Then Destino$ = SPY$
    '
    Call REPLA(X1$, UMED$, 73, 4)
    Call REPLA(X1$, MKD$(CDbl(CANORF)), 89, 8)
    Call REPLA(X1$, MKS$(LOSTA), 97, 4)
    Call REPLA(X1$, DESTI$, 101, 8)
    Call REPLA(X1$, MKI$(FENTREGA%), 109, 2)
    Call REPLA(X1$, SPX$, 115, 4)
    '
    Return
    '
VALIDA1:
    VALIDA% = (-1)
    X3$ = X2$: X1$ = X2$: X0$ = X2$
    If CVS(Mid$(X3$, 1, 4)) < 1 Then Return
    If CVI(Mid$(X3$, 69, 2)) < 1 Then Return
    If CVI(Mid$(X3$, 71, 2)) < 1 Then Return
    '
    VALIDA% = 0
    If TRIM$(Mid$(X3$, 73, 4)) = "" Then
        Call REPLA(X3$, "UNID", 73, 4)
    End If
    If CVD(Mid$(X3$, 89, 8)) < 1 Then Return
    If CVS(Mid$(X3$, 97, 4)) < 1 Then Return
    If CVI(Mid$(X3$, 109, 2)) < HO% Then
        Call COMUNI("Fecha de Entrega Proyectada\es Anterior a la Fecha Actual")
        Return
    End If
    FENTRE% = CVI(Mid$(X3$, 109, 2))
    '
    CAN# = CVD(Mid$(X3$, 89, 8))
    CANORF = CAN#
    If FRALOMAX% < 0 Then
       If CAN# > CVS(Mid$(X3$, 97, 4)) Then
          CALOTX = CAN#
          Call REPLA(X3$, MKS$(CALOTX), 97, 4)
       End If
    End If
    LOTE# = CDbl(CVS(Mid$(X3$, 97, 4)))
    If CAN# / LOTE# > 6 Then
        Call COMUNI("Cantidades no procesables\superan límite de 6 lotes\por Orden de Fabricación.")
        Return
    End If
    '
    DESTI$ = Mid$(X3$, 101, 8)
    DESTJ$ = ECOARCH$("DESTFAB", DESTI$)
    If TRIM$(DESTJ$) <> "" Then
        DESTI$ = DESTI$ + " - " + DESTJ$
    End If
    '
    If Asc(Mid$(X3$, 111, 1)) > 9 Then
      PRIO% = Asc(Mid$(X3$, 111, 1)) Mod 10
      Call REPLA(X3$, Chr$(PRIO%), 111, 1)
    End If
    PRIO% = Asc(Mid$(X3$, 111, 1))
    FLAN% = CVI(Mid$(X3$, 113, 2))
    Status% = 0
    If FLAN% <= HO% Then
       If HAYDISTO%(CIOF%, NRO&, CANORF) = 1 Then
          Status% = 1
       End If
    End If
    Call REPLA(X3$, Chr$(Status%), 112, 1)
    '
    CI% = CVI(Mid$(X3$, 71, 2))
    DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
    Call REPLA(X3$, DESCRI$, 5, 64)
    '
    TORFANUE$ = X3$     ' prepara registro para ORFANUE.DAT
    '
    FGEN = HO%
    STS$ = TRIM$(Str$(Status%)): While Len(STS$) < 3: STS$ = " " + STS$: Wend
    PRS$ = TRIM$(Str$(PRIO%)): While Len(PRS$) < 3: PRS$ = " " + PRS$: Wend
    SecProd = CVS(SPX$)
    FCIE = 0
    '
    X$ = REGBLAN$("INDIORF")
    Call REPLA(X$, MKS$(NRO&), 1, 4)
    Call REPLA(X$, FECHATEX$(FGEN), 5, 8)
    If Status% > 0 Then Call REPLA(X$, STS$, 13, 3)
    If PRIO% > 0 Then Call REPLA(X$, PRS$, 16, 3)
    Call REPLA(X$, ECOARCH$("SECPROD", MKS$(SecProd)), 21, 20)
    If FCIE <> 0 Then Call REPLA(X$, FECHATEX$(FCIE), 43, 8)
    '
    TINDIORF$ = X$      ' prepara registro para INDIORF.DAT
    VALIDA% = 1
    Return
    '
OBSERVOF:
    VALIDA% = 0
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    If REPIM% <> 2 Then
       If OBSERVOF% = 1 Then
          VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Fabricacion Numero" + Str$(CVS(Left$(X3$, 4))))
       End If
    End If
    Return
    '
GELOTES:
    NORDE% = 0
    CAPEN# = CAN#
    '
    While CAPEN# > 0
        CALOT# = LOTE#
        If CALOT# > CAPEN# Then
            CALOT# = CAPEN#
        End If
        '
        NORDE% = NORDE% + 1
        CILOTE%(NORDE%) = CI%
        CALOTE#(NORDE%) = CALOT#
        CAPEN# = CAPEN# - CALOT#
    Wend
    Return
    '
GESUBOR:
    If GESUBOR% = 1 Then
        CIXX% = CIOF%
        Call EXPLOMUL(CIXX%)
        For II% = 1 To CONTAMUL%
            PAD%(II%) = 1
        Next II%
        PAD%(CONTAMUL%) = 0
        '
        For II% = 1 To CONTAMUL% - 1
            If NIVE%(II% + 1) <= NIVE%(II%) Then
                PAD%(II%) = 0
              Else
                FIN& = ULTREG&("ARTIPRO")
                For J& = 1 To FIN&
                  X$ = REGLEIDO$("ARTIPRO", J&)
                    If CVI(Left$(X$, 2)) = CIK%(II%) Then
                      If CVI(Mid$(X$, 3, 2)) > 0 Then
                        PAD%(II%) = 0
                        GoTo 29
                      End If
                    End If
                Next J&
            End If
29      Next II%
        '
        For II% = 2 To CONTAMUL%
            If PAD%(II%) = 1 Then
                NORDE% = NORDE% + 1
                CILOTE%(NORDE%) = CIK%(II%)
                CALOTE#(NORDE%) = CAN# * CANT(II%)
            End If
        Next II%
    End If
    Return
    '
GERESMAT:
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    Call BLANARCH("!RESERVA")
    '
    For II% = 1 To NORDE%
       PUNMAT$(II%) = ""
       CODICONJ$ = CODEXT$(CILOTE%(II%))
       NORFABRI$ = ""
       CACONJU# = CALOTE#(II%)
       VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
       '
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 1)
       If VACONTROL% < 0 Then Return         ' cancelacion
       '
       HAYFALTA% = 0
       For U& = 1 To ULTREG&("!MATEROF")
         ZZ$ = REGLEIDO$("!MATEROF", U&)
         If CVS(Mid$(ZZ$, 25, 4)) > 0.05 Then HAYFALTA% = 1
       Next U&
       If HAYFALTA% > 0 Then
         If COMALTER%("Hay Faltantes de Materiales\para esta Orden de Fabricación.\\Cancelar\Generar Igual") <> 2 Then
           VACONTROL% = (-1)
           Return
         End If
       End If
       '
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       '
       
       For U& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U&)
          CIIX% = CVI(Left$(ZZ$, 2))
          CANTX# = CVS(Mid$(ZZ$, 5, 4))
          USOII = CANTX# / CACONJU#
          '
          If CANTX# >= 0.05 Then
             YY$ = REGBLAN$("RESERVA")
             Call REPLA(YY$, MKI$(CIIX%), 1, 2)
             Call REPLA(YY$, NORFA$, 3, 12)
             Call REPLA(YY$, MKI$(CIIX%), 15, 2)
             Call REPLA(YY$, MKI$(HO%), 17, 2)
             Call REPLA(YY$, MKD$(CANTX#), 19, 8)
             Call REPLA(YY$, MKD$(CANTX#), 35, 8)
             Call REPLA(YY$, MKS$(USOII), 75, 4)
             Call REGAPP("!RESERVA", YY$)
           PUNMAT$(II%) = PUNMAT$(II%) + MKS$(ULTREG&("!RESERVA"))
         End If
       Next U&
    Next II%
    '
    Call CIERRARCH("!RESERVA")
    Return
    '
GRAORFNUE:
    Call REPLA(TINDIORF$, MKS$(NRO&), 1, 4)
    '
    Z$ = RECFILT$("INDIORF", 1, MKS$(NRO&))
    If Len(Z$) >= 4 Then
          REGXX& = CVS(Left$(Z$, 4))
          Call GRAREG("INDIORF", TINDIORF$, REGXX&)
        Else
          Call REGAPP("INDIORF", TINDIORF$)
    End If
Return
'
GRAORDEFA:
    Call ABREORFAB
    '
    If REPIM% = 2 Then
      With OrdeFabr
        .MoveFirst
        Do While Not .EOF
          If !NUMEORFA = NRO& Then
            .Delete
          End If
        .MoveNext
        Loop
      End With
    End If
    '
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    For II% = 1 To NORDE%
      NORFA$ = NBRS$ + TRIM$(Str$(II%))
      DESCRI$ = TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT$(CILOTE%(II%)))
      '
      With OrdeFabr
        .AddNew
          !NUMEORFA = NRO&
          !Referen = AJUSTI$(DESCRI$, 64)
          !IdSubOr = NORFA$
          !FechGen = Now
          !Cod_Inte = CILOTE%(II%)
          !COD_EXTE = CODEXT$(CILOTE%(II%))
          !Descrip = DESCRIT0$(CILOTE%(II%))
          !Unidad = UNIMED$(CILOTE%(II%))
          !CanProy = CALOTE#(II%)
          If Len(DESTI$) > 8 Then
          DESTI$ = TRIM$(Left(DESTI$, 8))
          End If
          !Destino = DESTI$
          !FechEnSol = CVDAT(FENTRE%)
          !Priorid = PRIO%
          !Statuorf = Status%
          !FechLan = CVDAT(FLAN%)
          !SecciPro = SecProd
          !Fechcierr = CVDAT(0)
          !CanApro = 0
          !CanRech = 0
          !IdenRetra = ""
          !FechProyEntre = CVDAT(0)
          !FechTempEntre = CVDAT(0)
          !FechRecal = CVDAT(0)
          !Observa = ""
        .Update
      End With
      '
      CANS$ = " - (" + TRIM$(Str$(CALOTE#(II%))) + ")"
      DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
      ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
      RFG$ = RECFILT$("ECORDEF", 1, NORFA$)
      If Len(RFG$) >= 4 Then
          REODE& = CVS(Left$(RFG$, 4))
          Call GRAREG("ECORDEF", ECORDEF$, REODE&)
        Else
          Call REGAPP("ECORDEF", ECORDEF$)
      End If
      '
      If Status% = 1 Then
        RFG$ = RECFILT$("ECORDEP", 1, NORFA$)
        If Len(RFG$) >= 4 Then
            REODE& = CVS(Left$(RFG$, 4))
            Call GRAREG("ECORDEP", ECORDEF$, REODE&)
          Else
            Call REGAPP("ECORDEP", ECORDEF$)
        End If
      End If
      '
    Next II%
Return
    '
GRARESERV:
    Call ABRESERVA
    'LONUO% = Len(NBRS$)
    For U& = 1 To ULTREG&("!RESERVA")
      YY$ = REGLEIDO$("!RESERVA", U&)
      'Call REPLA(YY$, NBRS$, 3, LONUO%)
      NORFX$ = TRIM$(Mid$(YY$, 3, 12))
      Call GRAREG("!RESERVA", YY$, U&)
      '
      CIRF% = CVI(Left$(YY$, 2))
      CANTRX# = CVD(Mid$(YY$, 19, 8))
      USOIX = CVS(Mid$(YY$, 75, 4))
      IDLOT$ = TRIM$(Mid$(YY$, 43, 16))
      DEPOX% = Asc(Mid$(YY$, 59, 1))
      '
      If REPIM% = 2 Then
        With Reserva
          .MoveFirst
          Do While Not .EOF
            If !IdSubOr = NORFX$ Then
              .Delete
            End If
          .MoveNext
          Loop
        End With
      End If
      '
      With Reserva
        .AddNew
        !Cod_Inte = CIRF%
        !COD_EXTE = CODEXT$(CIRF%)
        !Descrip = DESCRIT0$(CIRF%)
        !IdSubOr = NORFX$
        !FechRes = Now
        !CantRes = CANTRX#
        !CantCons = 0
        !Saldo_Entre = CANTRX#
        !IdenLote = IDLOT$
        !Deposito = DEPOX%
        !UsoUnit = USOIX
        .Update
      End With
    Next U&
    Reserva.Close
Return
    '
DEPUORFEM:
    'If MODORFA% = 2 Then
    Call BLOQARCH("ORFAEMI")
    '
    For II% = 1 To NORDE%
      CALOTEX# = CALOTE#(II%)
      CIFX% = CILOTE%(II%)
      If CIFX% > 0 Then
        RFOE$ = RECFILT$("ORFAEMI", 1, MKI$(CIFX%))
        For KKU& = 1 To Len(RFOE$) Step 4
          REOFEM& = CVS(Mid$(RFOE$, KKU&, 4))
          X$ = REGLEIDO$("ORFAEMI", REOFEM&)
          CAMENO# = CVD(Mid$(X$, 49, 8))
          If CAMENO# <= CALOTEX# Then
              CALOTEX# = CALOTEX# - CAMENO#
              CAMENO# = 0
            Else
              CAMENO# = CAMENO# - CALOTEX#
              CALOTEX# = 0
          End If
          Call REPLA(X$, MKD$(CAMENO#), 49, 8)
          Call REPLA(X$, FORMATNUM$(CAMENO#, "F10.1"), 29, 10)
          Call GRAREG("ORFAEMI", X$, REOFEM&)
          If CALOTEX# < 0.05 Then Exit For
        Next KKU&
      End If
    Next II%
    '
    JK& = 0: FIN& = ULTREG&("ORFAEMI")
    For REOFEM& = 1 To FIN&
      X$ = REGLEIDO$("ORFAEMI", REOFEM&)
      If CVD(Mid$(X$, 49, 8)) >= 0.05 Then
        JK& = JK& + 1
        Call GRAREG("ORFAEMI", X$, JK&)
      End If
    Next REOFEM&
    Call SETULTREG&("ORFAEMI", JK&)
    Call CIERRARCH("ORFAEMI")
    Call FRESHECHO("ORFAEMI")
    '
    Return
    '
GRASECOP:
    Call BLANARCH("!OPERFAB")
    Call ABREORFAB
    '
    For II% = 1 To NORDE%
      NORFA$ = NBRS$ + TRIM$(Str$(II%))
      CIJK% = CILOTE%(II%)
      CAPENDE# = CALOTE#(II%)
      '
      If REPIM% = 2 Then
        With OperFab
1221      .FindFirst "IdSubOr = '" & NORFA$ & "'"
          If .NoMatch = False Then
            .Delete
            GoTo 1221
          End If
        End With
      End If
      '
      Set SecopTemp = Articulos.OpenRecordset("SELECT * FROM SecOper WHERE CodIConj=" & CIJK% & " AND StatOper = 1 ORDER BY NumeOper ASC", dbOpenSnapshot)
      On Error Resume Next
      SecopTemp.MoveFirst
      If Err < 1 Then
        With SecopTemp
          Do While Not .EOF
            NO% = !NumeOper
            DOPE$ = !DescOper
            EQ1$ = !CodMaq0
            EQAUX$ = AJUSTI$(!CodMaq1, 6) + AJUSTI$(!CodMaq2, 6) + AJUSTI$(!CodMaq1, 6)
            LOTESTA = !LOTESTAN: If LOTESTA < 1 Then LOTESTA = 1
            CALOTES = 1
            CAOPS = !CantiOps
              If CAOPS < 0.01 Then CAOPS = 1
            CICLOHE = !PieCiclo
              If CICLOHE < 1 Then CICLOHE = 1
            MINPREP = 60 * !HorsPrep * CALOTES
              If modo% > 0 Then MINPREP = 0
            MINFIJO = 60 * !HorsFijo * CALOTES
            MINVARI = CAPENDE# * !MinuStan
            MINHERR = CAPENDE# * 60 * !HorsMHer / CICLOHE
            MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
            MINLIMP = 60 * !HorsLimp * CALOTES
            MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO + MINLIMP
            HORAPROD = MINLOTE / 60
            '
            BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
            NUOROP% = 0
            FEAL% = HOY(HOS$)
            HOAL% = HORANUM(Time$)
            USUA% = USNBR% Mod 100
            DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
            INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
            PROGBLA$ = Space$(6) + String$(26, 0)
            '
            'If REPIM% <> 2 Then
              If Status% = 1 Then
                With OperFab
                  .AddNew
                    !IdSubOr = NORFA$
                    !Cod_Inte = CIJK%
                    !COD_EXTE = CODEXT$(CIJK%)
                    !Despcrip = DESCRIT0$(CIJK%)
                    !NumeOper = NO%
                    !DescOper = DOPE$
                    !Maquina = EQ1$
                    !DescMaq = ECOARCH("PADMAQ", EQ1$)
                    !EquiAux = EQAUX$
                    !CantiOps = CAOPS
                    !CantProg = CAPENDE#
                    !CantReal = 0
                    !CantPend = CAPENDE#
                    !TiemPrep = MINPREP
                    !TiemFijo = MINFIJO
                    !TiemVari = MINVARI
                    !TiemDemo = MINDEMO
                    !TiemMHer = MINHERR
                    !TiemLimp = MINLIMP
                    !TiemOtro = 0
                    !TiemTota = MINLOTE
                    !TiemEjec = 0
                    !TiemPend = MINLOTE
                    !StatOper = 1
                    !ComiePro = CVDAT(0)
                    !TermiPro = CVDAT(0)
                    !TermiTemp = CVDAT(0)
                    !Observa = " "
                  .Update
                End With
              End If
            'End If
          .MoveNext
          Loop
        End With
      End If
    Next II%
    '
    OperFab.Close
    Return
    '
PRINORFAB:
237 If TRIM$(FORIPRE$) = "" Then Return
    '
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    DETDES = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
    DETHAS = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
    INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
    Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
       INTERLI = 4
    End If
    '
    MAXITEMS = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXITEMS < 1 Then MAXITEMS = 999
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
       CACOPI% = 1
    End If
    '
    PORIAN = Printer.Orientation  ' retiene orientacion
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    ORIENTACT = 1       ' anterior
    NORPAPSI = Printer.PaperSize
    '
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
       ORIENTACT = 2
    End If
    '
    DESCRIOR$ = TRIM$(CODEXT$(CILOTE%(1))) + " - " + TRIM$(DESCRIT0$(CILOTE%(1)))
    '
    For II% = 1 To NORDE%
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KKINI% = 1
249       KK1% = 1: NH% = 1
250       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          Call GENEOPERORF(NORFA$)
          '
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          Call PRINTFORWIN(DESCRIOR$, "REF-MAT1", 0, 0)
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
          '
          Call PRINTFORWIN(ECOARCH$("SECPROD", SPX$), "REF-MAT3", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(II%))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(II%)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(FEPLANO$(CILOTE%(II%)), "F-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(II%)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(II%)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(II%)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(II%)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(II%))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%) * PESUNI(CILOTE%(II%))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMPX$), "REF-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(LARCUTE$), "LAR-CUTE", 0, 0)
          Call PRINTFORWIN(TRIM$(ACHCUTE$), "ACH-CUTE", 0, 0)
          Call PRINTFORWIN(TRIM$(LARMATE$), "LAR-MATE", 0, 0)
          Call PRINTFORWIN(TRIM$(ACHMATE$), "ACH-MATE", 0, 0)
          Call PRINTFORWIN(TRIM$(LARCUCO$), "LAR-CUCO", 0, 0)
          Call PRINTFORWIN(TRIM$(ACHCUCO$), "ACH-CUCO", 0, 0)
          Call PRINTFORWIN(TRIM$(LARMACO$), "LAR-MACO", 0, 0)
          Call PRINTFORWIN(TRIM$(ACHMACO$), "ACH-MACO", 0, 0)
          '
          YAC = DETDES
          ITNU% = 0
          If DETHAS > DETDES + INTERLI Then
            '
            EPAC$ = TRIM$(UCase$(EMPREAC$))
            If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
              If Len(PUNMAT$(II%)) > 4 Then
                Call PRINTFORWIN("Segun Vale de Materiales", "DES-MPRI", 0, YAC)
                GoTo 260
              End If
            End If
            '
            For KK% = KK1% To Len(PUNMAT$(II%)) Step 4
              TTL& = CVS(Mid$(PUNMAT$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 260
                '
                If YAC + INTERLI > DETHAS Then
                   KK1% = KK%
                   ATTX% = 1
                   NH% = NH% + 1
                   GoTo 260
                End If
                '
                YY$ = REGLEIDO$("!RESERVA", TTL&)
                CIXI% = CVI(Mid$(YY$, 1, 2))
                CANTX# = CVD(Mid$(YY$, 19, 8))
                '
                If CANTX# > 0 Then
                  Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                  Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                  Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                    STACTU = STOCKACT(CIXI%): If STACTU < 0 Then STACTU = 0
                  Call PRINTFORWIN(FORMATNUM$(STACTU, "F12.1"), "CANT-DIS", 0, YAC)
                  Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                  Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                  Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                  Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                  Call PRINTFORWIN(PARTHIL$, "PAR-MPRI", 0, YAC)
                    CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                  Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                    CANS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                  Call PRINTFORWIN(CANS$, "USO-MPRI", 0, YAC)
                  Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                  YAC = YAC + INTERLI
                End If
              End If
            Next KK%
            '
260         YAC = DETDES
            ITNU% = 0
            TMPRE = 0: TMPRO = 0
            If KKINI% < 1 Then KKINI% = 1
            For KK% = 1 To ULTREG&("!OPERFAB")    'KKINI% To Len(PUNOPE$(II%)) Step 4
              TTL& = KK%   '     CVS(Mid$(PUNOPE$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
                  If YAC + INTERLI > DETHAS Then
                    KKINI% = KK%
                    Call SETSPOOL("NEWPAGE", "")
                    GoTo 249
                  End If
                End If
                '
                YY$ = REGLEIDO$("!OPERFAB", TTL&)
                '
                NOPE% = CVI(Mid$(YY$, 15, 2))
                DOPE$ = Mid$(YY$, 17, 48)
                DESOP$ = OBSEROP$(CILOTE%(II%), NOPE%)
                If TRIM$(DESOP$) = "" Then DESOP$ = DOPE$
                EQ1$ = Mid$(YY$, 73, 6)
                EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
                EQ2$ = Mid$(YY$, 203, 6)
                EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
                EQ3$ = Mid$(YY$, 209, 6)
                EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
                MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
                MPRO = 0: For KI% = 107 To 123 Step 4
                              MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                           Next KI%
                CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
                TMPRE = TMPRE + MPRE
                TMPRO = TMPRO + MPRO
                MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
                MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 269
                '
                Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DESOP$), "OBS-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
                Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
                Call PRINTFORWIN(NUHOPROF$(CILOTE%(II%), NOPE%), "N-HPROF", 0, YAC)
                '
                YAC = YAC + INTERLI
                '
              End If
269         Next KK%
            '
            TMPREI% = 0: If TMPRE < 16384 Then TMPREI% = TMPRE
            TMPROI% = 0: If TMPRO < 16384 Then TMPROI% = TMPRO
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI%), 6), "TOT-PREP", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPROI%), 6), "TOT-PROD", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI% + TMPROI%), 6), "TOT-TIEM", 0, 0)
            '
          End If
          '
          GoSub PRIPLACONTROL
          GoSub PRINPUEPUNTO
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
   
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
          '
          If ATTX% <> 0 Then GoTo 250
          '
       Next COP%
       '
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
       '
    Next II%
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    For II% = 1 To NORDE%
       '
       CIXOFX% = CILOTE%(II%)
       PUNOPFABRI$ = PUNOPE$(II%)
       GoSub DOCUANEXOS
       '
    Next II%
    '
337 If TRIM$(FORIPRE2$) = "" Then GoTo 990
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
      If Len(PUNMAT$(1)) <= 4 Then
        GoTo 990
      End If
    End If
    '
    PORIAN = Printer.Orientation
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    CODFOR$ = TRIM$(UCase$(FORIPRE2$))
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    DETDES = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
    DETHAS = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
    INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
    Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
       INTERLI = 4
    End If
    '
    MAXITEMS = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXITEMS < 1 Then MAXITEMS = 999
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
       CACOPI% = 1
    End If
    '
    ORIENTACT = 1
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
       ORIENTACT = 2
    End If
    '
    If VALINDOP% = 1 Then GoTo 445 ' caso vales de MAINIERI
    '
345 For II% = 1 To NORDE%
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KK1% = 1: NH% = 1
350       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          Call GENEOPERORF(NORFA$)
          '
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(II%))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(II%)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(II%)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(II%)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(II%)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(II%)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(II%))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%) * PESUNI(CILOTE%(II%))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMPX$), "REF-CONJ", 0, 0)
          '
          YAC = DETDES
          ITNU% = 0
          If DETHAS > DETDES + INTERLI Then
            For KK% = KK1% To Len(PUNMAT$(II%)) Step 4
              TTL& = CVS(Mid$(PUNMAT$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 360
                '
                If YAC + INTERLI > DETHAS Then
                   KK1% = KK%
                   ATTX% = 1
                   NH% = NH% + 1
                   GoTo 360
                End If
                '
                YY$ = REGLEIDO$("!RESERVA", TTL&)
                CIXI% = CVI(Mid$(YY$, 1, 2))
                CANTX# = CVD(Mid$(YY$, 19, 8))
                '
                Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                Call PRINTFORWIN(PARTHIL$, "PAR-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                Call PRINTFORWIN(CANS$, "USO-MPRI", 0, YAC)
                Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                YAC = YAC + INTERLI
              End If
            Next KK%
            '
360         YAC = DETDES
            ITNU% = 0
            TMPRE = 0: TMPRO = 0
            For KK% = 1 To ULTREG&("!OPERFAB")     ' To Len(PUNOPE$(II%)) Step 4
              TTL& = KK%                           ' CVS(Mid$(PUNOPE$(II%), KK%, 4))
              If TTL& > 0 Then
               '
               YY$ = REGLEIDO$("!OPERFAB", TTL&)
                '
                NOPE% = CVI(Mid$(YY$, 15, 2))
                DOPE$ = Mid$(YY$, 17, 48)
                EQ1$ = Mid$(YY$, 73, 6)
                EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
                EQ2$ = Mid$(YY$, 203, 6)
                EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
                EQ3$ = Mid$(YY$, 209, 6)
                EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
                MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
                MPRO = 0: For KI% = 107 To 123 Step 4
                              MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                           Next KI%
                CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
                TMPRE = TMPRE + MPRE
                TMPRO = TMPRO + MPRO
                MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
                MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 369
                '
                Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
                Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
                '
                YAC = YAC + INTERLI
                '
              End If
369         Next KK%
            '
            TMPREI% = 0: If TMPRE < 16384 Then TMPREI% = TMPRE
            TMPROI% = 0: If TMPRO < 16384 Then TMPROI% = TMPRO
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI%), 6), "TOT-PREP", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPROI%), 6), "TOT-PROD", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI% + TMPROI%), 6), "TOT-TIEM", 0, 0)
            '
          End If
          '
370       YAC = DETDES
          ITNU% = 0
          For TTL& = 1 To ULTREG&("!BCOMPLEM")
             '
             YY$ = REGLEIDO$("!BCOMPLEM", TTL&)
             '
             ITNU% = ITNU% + 1
             If ITNU% > MAXITEMS Then GoTo 379
             '
             TTXX$ = Mid$(YY$, 3, 14)
             Call PRINTFORWIN(TRIM$(TTXX$), "PIE-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 17, 16)
             Call PRINTFORWIN(TRIM$(TTXX$), "DIB-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 33, 6)
             Call PRINTFORWIN(TRIM$(TTXX$), "MAQ-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 39, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "LAR-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 43, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "ACH-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 47, 2)
             Call PRINTFORWIN(TRIM$(TTXX$), "UNI-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 53, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "CAN-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 61, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "PES-COMP", 0, YAC)
             '
             YAC = YAC + INTERLI
379       Next TTL&
          '
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
          '
          If ATTX% <> 0 Then GoTo 350
          '
       Next COP%
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Orden Complem. Fabricación " + NORFA$, OKX%)
       '
    Next II%
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    GoTo 990
    
    '
445 ' CASO VALES DE MAINIERI
    PORIAN = Printer.Orientation
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    For II% = 1 To ULTREG&("!OPERFAB")        ' Len(PUNOPE$(1)) / 4
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KK1% = 1: NH% = 1
450       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(1))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(1)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(1)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(1)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(1)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(1)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(1))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(1)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(1) * PESUNI(CILOTE%(1))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMPX$), "REF-CONJ", 0, 0)
          '
460       YAC = DETDES
          ITNU% = 0
          TMPRE = 0: TMPRO = 0
          TTL& = II%        '   CVS(Mid$(PUNOPE$(1), 4 * II% - 3, 4))
          '
          If TTL& > 0 Then
             '
             YY$ = REGLEIDO$("!OPERFAB", TTL&)
             '
             NOPE% = CVI(Mid$(YY$, 15, 2))
             DOPE$ = Mid$(YY$, 17, 48)
             EQ1$ = Mid$(YY$, 73, 6)
             EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
             EQ2$ = Mid$(YY$, 203, 6)
             EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
             EQ3$ = Mid$(YY$, 209, 6)
             EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
             MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
             MPRO = 0: For KI% = 107 To 123 Step 4
                           MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                        Next KI%
             CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
             TMPRE = TMPRE + MPRE
             TMPRO = TMPRO + MPRO
             MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
             MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
             '
             Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
             Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
             Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
             '
             YAC = YAC + INTERLI
             '
          End If
          '
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
       Next COP%
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Vales de Operación O/F " + NORFA$, OKX%)
       '
    Next II%
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    GoTo 990
    '
    '
    '
990 'Printer.Orientation = PORIAN ' restaura anterior
    'Call SETSPOOL("ORIENTATION", Str$(PORIAN))
    'Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
    If CANCELPRINT% < 1 Then
      If COMALTER%("Desea Continuar o Repetir\la Impresion de este Documento ?\\Continuar\Repetir Impresion") = 2 Then GoTo 237
    End If
    '
    DOCUORIG$ = TRIM$(NORFA$)
    NUMEDOC$ = ""
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ARCON") > 0 Then
       '
       Call CIERRARCH("*.*")
       Call BAJALDISCO
       '
       ANEXOCOM.Caption = "Ordenes de Fabricación - Documentos Anexos"
       YY$ = REGBLAN$("OCOMDETA")
       Call REPLA(YY$, MKI$(CILOTE%(1)), 9, 2)
       Call GRAREG("!OCOMDETA", YY$, 1)
       Call SETULTREG("!OCOMDETA", 1)
       Call CIERRARCH("!OCOMDETA")
       '
       ANEXOCOM.Show 1
       '
    End If
    '
Return
    '
    '
PRIPLACONTROL:
    '
    YAC = DETDES
    PPCX$ = RECFILT$("PLACONTR", 1, MKI$(CIOF%))
    INTERLIX = INTERLI
    '
    If Len(PPCX$) >= 4 Then
       EPAC$ = EMPREAC$
       If InStr(EPAC$, "NEUMANN") > 0 Then
          CARENEC% = Len(PPCX$) / 4
          INTERLI2 = (190 - 152) / (CARENEC% + 1)
          If INTERLI2 < INTERLIX Then INTERLIX = INTERLI2
       End If
    End If
    '
    For KK% = 1 To Len(PPCX$) Step 4
       TTL& = CVS(Mid$(PPCX$, KK%, 4))
       If TTL& > 0 Then
          If TTL& <= ULTREG&("PLACONTR") Then
             YY$ = REGLEIDO$("PLACONTR", TTL&)
             Call PRINTFORWIN(Mid$(YY$, 3, 2), "COT-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 5, 28), "REF-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 33, 32), "MED-PLCO", 0, YAC)
             Call PRINTFORWIN(Str$(CVS(Mid$(YY$, 65, 4))), "CAN-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 69, 12), "FRE-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 81, 160), "REA-PLCO", 0, YAC)
             '
             YAC = YAC + INTERLIX
          End If
       End If
    Next KK%
    '
    Return
    '
    '
PRINPUEPUNTO:
    '
    YAC = DETDES
    PPCX$ = RECFILT$("PUEPUNTO", 1, MKI$(CIOF%))
    For KK% = 1 To Len(PPCX$) Step 4
       TTL& = CVS(Mid$(PPCX$, KK%, 4))
       If TTL& > 0 Then
          If TTL& <= ULTREG&("PUEPUNTO") Then
             YY$ = REGLEIDO$("PUEPUNTO", TTL&)
             Call PRINTFORWIN(Mid$(YY$, 3, 24), "OPE-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 27, 24), "MAQ-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 51, 24), "PAR-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 75, 24), "HER-PUEP", 0, YAC)
             '
             YAC = YAC + INTERLI
          End If
       End If
    Next KK%
    '
Return
    '
    '
DOCUANEXOS:
    '
    DOCUORIG$ = NORFA$
    '
    If CONTROL$("ORDEFABR", "PRIPLANO") = "SI" Then
       Call PRIPLANO(CIXOFX%)
    End If
    '
    If CONTROL$("ORDEFABR", "PRIESPEC") = "SI" Then
       Call PRIHOJESP(CIXOFX%)
    End If
    '
    If CONTROL$("ORDEFABR", "PRIHOPRO") = "SI" Then
       For KK% = 1 To Len(PUNOPFABRI$) Step 4
          TTL& = CVS(Mid$(PUNOPFABRI$, KK%, 4))
          If TTL& > 0 Then
            YY$ = REGLEIDO$("!OPERFAB", TTL&)
            NOPE% = CVI(Mid$(YY$, 15, 2))
            Call PRIHOJPROFAB(CIXOFX%, NOPE%)
          End If
       Next KK%
    End If
    '
    If CONTROL$("ORDEFABR", "PRIHOINS") = "SI" Then
       Call PRIHOJRINSP(CIXOFX%)
    End If
    '
    DOCUORIG$ = ""
    '
Return
    '
    '
999 Call CIERRARCH("*.*")
    Call FRESHALL
End Sub
'
Private Sub GenReservOC_Click()
Call RESERV04.Command3_Click
End Sub

Private Sub InfoPro_Click()
Call Command51_Click
End Sub

Private Sub LanOrd_Click()
Call Command42_Click
End Sub

Private Sub LiToImpre_Click()
'
RESERV04.Option6 = True
RESERV04.Option4 = True
Call RESERV04.Command2_Click
'

End Sub

Private Sub LiToPan_Click()
'
RESERV04.Option6 = True
RESERV04.Option5 = True
Call RESERV04.Command2_Click
'
End Sub



Private Sub MaestArt_Click()
   '
       Call CONECRUN("ARCHIG04/AMASTO", "Padron Maestro de Articulos")
   '
End Sub

Private Sub ModiOrd_Click()
Call Command43_Click
End Sub

Private Sub ModOrd_Click()
Call Command43_Click
End Sub

Private Sub NvaOrd_Click()
Call Command41_Click
End Sub

Private Sub OperPro_Click()
Call Command7_Click
End Sub



Private Sub otraccdir_Click()
Call Command8_Click
End Sub

Private Sub PenNumOF_Click()
Option1.Value = True
Call Command63_Click
End Sub

Private Sub porPantall_Click()
Option1.Value = True
Call Command64_Click
End Sub

Private Sub ReXMat_Click()
'
RESERV04.Option3 = True
RESERV04.Option5 = True
Call RESERV04.Command2_Click
'

End Sub

Private Sub ReXNumOF_Click()
'
RESERV04.Option2 = True
RESERV04.Option5 = True
Call RESERV04.Command2_Click
'

End Sub

Private Sub SuOrd_Click()
Call Command44_Click
End Sub

Private Sub SusOrd_Click()
Call Command44_Click
End Sub



Private Sub ValPrep_Click()
Call VALEMATER(1)
End Sub

Private Sub ValTran_Click()
Call VALEMATER(2)
End Sub

Private Sub VerImp_Click()
Call Command5_Click
End Sub

Private Sub VeryReimp_Click()
Call Command5_Click
End Sub

Private Sub VeryReimpi_Click()
Call Command5_Click
End Sub

Private Sub Ximpre_Click()
Option2.Value = True
Call Command62_Click
End Sub

Private Sub Ximpres_Click()
Option2.Value = True
Call Command10_Click
End Sub

Private Sub Ximpreso_Click()
Option2.Value = True
Call Command64_Click
End Sub

Private Sub Xpant_Click()
Option1.Value = True
Call Command62_Click
End Sub

Private Sub Xpanta_Click()
Option1.Value = True
Call Command10_Click
End Sub

Sub REPRINORFA()
   '
   Call ORFNUE(2)
   '
End Sub
'
Sub REAPEROF()
    '
    Call ABREORFAB
    SQLX$ = " SELECT DISTINCT IdSubOr, StatuOrf, CanApro, CanProy FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE StatuOrf = 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "No Existen Ordenes de Fabricación\que Puedan Re-Abrirse.")
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    '
90  Call SETULTREG("!ECORDEP", JJ&)
    Call CIERRARCH("!ECORDEP")
    Call FRESHECHO("!ECORDEP")
    Screen.MousePointer = 1
    '
    Call SELECHO("!ECORDEP")
    NOFA$ = TRIM$(VALACT1$("!ECORDEP"))
    If NOFA$ <> "" Then
      If COMALTER%("Confirme la Re-Apertura de la " + NOFA$ + "\\Si, Re-Abrir\No, Cancelar") = 1 Then
        Call ABREORFAB
        PUDOABRIRSE% = 0
        With OrdeFabr
          .FindFirst "Idsubor = '" & NOFA & "' AND StatuOrf = 3 "
          If .NoMatch = False Then
            .Edit
            !Statuorf = 1
            .Update
            PUDOABRIRSE% = 1
            .FindNext "Idsubor = '" & NOFA & "' AND StatuOrf = 3 "
          End If
        End With
        If PUDOABRIRSE% = 0 Then
            Call COMUNI("No Fue Posible Re-Abrir esta O/F")
          ElseIf PUDOABRIRSE% = 1 Then
            Call COMUNI(NOFA$ + " Re-Abierta.\No se Re-Generaron Reservas ni Operaciones")
        End If
      End If
    End If
    '
99  OrdeFabr.Close
    '
End Sub
Sub ENVPREN()
    '
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "ENVIOS DE PRENDA - GENERALES"

100  NROOF$ = ""
     '
     Call GENRESDAT
     '
1405 Call FRESHECHO("ECORDEP")
     If ULTREG&("ECORDEP") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
         Exit Sub
     End If
     '
     Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
     If TRIM$(NROOF$) = "" Then
       Exit Sub
     End If
     '
     Screen.MousePointer = 11
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM OPERFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     SQLX$ = SQLX$ + " ORDER BY NumeOper "
     '
     Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esta Orden de Fabricación.")
         On Error GoTo 0
         GoTo 1405
       End If
       On Error GoTo 0
       '
       FORSELOPE.List1.Clear
       Do While (.EOF = False)
          CAPROD# = !CantProg      'Cantidad Programada
          CAREA# = !CantReal       'Cantidad realizada
          SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
          '
          If SALDO# > 0.05 Then
            NOPE% = !NumeOper      'numero de operación
            DOPE$ = !DescOper      'Descripcion de la operacion
            MAQUI$ = !Maquina      'Máquina
            '
            JJ& = JJ& + 1: JJI% = JJ&
            YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
            NOPX$ = AJUSTI$(TRIM$(Str$(NOPE%)) + " - " + TRIM$(Left$(DOPE$, 31)) + " - " + MAQUI$, 46)
            'NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            FORSELOPE.Caption = "Selección de Operaciones de: " & NROOF$
            FORSELOPE.LIMATSEL.AddItem TRIM(YY$)
            Screen.MousePointer = 1
          End If
       .MoveNext
       Loop
       FORSELOPE.Show 1
     End With
     Screen.MousePointer = 11
     '
     With FORSELOPE
       CANTI1 = 0
       For JJ& = 0 To .LIMATSEL.ListCount - 1
         Screen.MousePointer = 11
         Y$ = .LIMATSEL.List(JJ&)
         CANTI = Val(Mid$(Y$, 35, 50))
         If CANTI > CANTI1 Then
           CANTI1 = CANTI 'ME QUEDO CON LA CANT MENOR
         End If
         Screen.MousePointer = 1
       Next JJ&
     End With
     Unload FORSELOPE
     'CON ESTE CODIGO LLAMO AL OBJETO
     'SACO EL PRODUCTO DE LA O/F
     SQLX$ = " IdSubOr = '" & NROOF$ & "' "
     With OrdeFabr
      CI% = 0
      .FindFirst SQLX$
      If .NoMatch = False Then
        CI% = !Cod_Inte
      End If
     End With
     '
1415 YY$ = ZZ$ + Space$(320)   ' espacios para observaciones
     Call REPLA(YY$, NROOF$, 1, 12) 'NRO OF
     Call REPLA(YY$, MKI(CI%), 13, 2) 'NRO INT MATERIAL
     Call REPLA(YY$, MKI$(0), 73, 2) ' NUMERO DE TALLER
     Call REPLA(YY$, MKS$(0), 75, 4) ' NUMERO DE REMITO
     Call REPLA(YY$, String$(178, 0), 79, 178)
     Call REPLA(YY$, MKD$(CAINF#), 79, 8) ' CANTIDAD PROYECTADA
     Call REPLA(YY$, MKI$(HO%), 135, 2)
     Call REPLA(YY$, MKI$(HO%), 139, 2)
     Call REPLA(YY$, Space$(30), 151, 30)
     Call REPLA(YY$, String$(76, 0), 181, 76)
     '
1420 XX$ = OBJPLA$("TERCERCONF", YY$)
     If XX$ = "" Then GoTo 1415
     '
     CAINF# = CVD(Mid$(XX$, 79, 8))
     If CAINF# < 0.999 Then
       Call MENSERR(24, "Falta Cantidad")
       YY$ = XX$
       GoTo 1420
     End If
     Call REPLA(XX$, MKD$(CAINF#), 95, 8)
     '
     NTALL% = CVI(Mid$(XX$, 73, 2))
     If NTALL% < 1 Then
        Call MENSERR(24, "Falta Taller")
        YY$ = XX$
        GoTo 1420
      ElseIf ECOARCH$("ACREDOR", MKI$(NPRO%)) = "" Then
        Call MENSERR(24, "Número de Taller no Válido")
        YY$ = XX$
        GoTo 1420
     End If
     FENTRESOL% = CVI(Mid$(XX$, 139, 2))
     Observa$ = TRIM(Mid$(XX$, 140, 300))
     '
     Screen.MousePointer = 1
     '
End Sub
Sub GENRESDAT()
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        Call REPLA(Y$, MKI$(!Cod_Inte), 1, 2)
        Call REPLA(Y$, !IdSubOr, 3, 12)
        Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(!CantRes), 19, 8)
        Call REPLA(Y$, MKD$(!CantCons), 27, 8)
        Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
        Call REPLA(Y$, !IdenLote, 43, 16)
        Call REPLA(Y$, Chr$(!Deposito), 60, 1)
        Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("RESERVA", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("RESERVA", JJ&)
    Call CIERRARCH("RESERVA")
    '
    Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, "NO Existen Ordenes de Fabricación. Imposible Continuar la Carga.")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
End Sub
