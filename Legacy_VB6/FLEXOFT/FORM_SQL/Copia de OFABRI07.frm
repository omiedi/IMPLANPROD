VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OFABRI07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O/F's - Ordenes de Fabricación"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   10770
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "OFABRI07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command55 
      Caption         =   "Recup.O.F. pdoc_tbl"
      Height          =   495
      Left            =   4680
      TabIndex        =   102
      ToolTipText     =   "Recuperacion de orden de fabricacion en Pdoc_tbl"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   100
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Recu FIFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   98
      ToolTipText     =   "Recupera Movimientos por Lotes"
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Recupera Sector Productivo"
      Height          =   225
      Left            =   6510
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   3060
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
      Picture         =   "OFABRI07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   "Listado Fabricación por Rango de Fechas"
      Top             =   7245
      Width           =   650
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   6840
      Left            =   9870
      ScaleHeight     =   6780
      ScaleWidth      =   1005
      TabIndex        =   87
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "OFABRI07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command53 
         Caption         =   "Of-Ns"
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
         Left            =   120
         TabIndex        =   97
         ToolTipText     =   "Ordenes de Fabricacion vs numeros de serie"
         Top             =   2160
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   91
         Top             =   3570
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "OFABRI07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Left            =   120
         Picture         =   "OFABRI07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1500
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   120
         Picture         =   "OFABRI07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2820
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OFABRI07.frx":11B4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3885
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
      TabIndex        =   70
      Top             =   6195
      Width           =   4425
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   81
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
            Picture         =   "OFABRI07.frx":30D6
            Style           =   1  'Graphical
            TabIndex        =   82
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
            TabIndex        =   83
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
         TabIndex        =   74
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
         TabIndex        =   71
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
            Picture         =   "OFABRI07.frx":33E0
            Style           =   1  'Graphical
            TabIndex        =   72
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
            TabIndex        =   73
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
      TabIndex        =   63
      Top             =   7455
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Asignación de Materias Primas"
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
      Left            =   180
      TabIndex        =   26
      Top             =   1575
      Width           =   4955
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2600
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   60
         Top             =   420
         Width           =   2000
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
            Picture         =   "OFABRI07.frx":36EA
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Transferencia Materiales de Depósito a Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia de Materiales"
            BeginProperty Font 
               Name            =   "Arial"
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
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   340
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   34
         Top             =   420
         Width           =   2000
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
            Picture         =   "OFABRI07.frx":3834
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Control y Gestión de Reservas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales Reservados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   735
            TabIndex        =   36
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
         TabIndex        =   64
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
            Picture         =   "OFABRI07.frx":3B3E
            Style           =   1  'Graphical
            TabIndex        =   69
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
            TabIndex        =   65
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
         TabIndex        =   66
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
            Picture         =   "OFABRI07.frx":3E48
            Style           =   1  'Graphical
            TabIndex        =   68
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
            TabIndex        =   67
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
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programación de Carga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5280
      TabIndex        =   10
      Top             =   2940
      Width           =   4425
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2310
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   54
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
            Picture         =   "OFABRI07.frx":4152
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Consulta de Tiempo Total Programado por Equipo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga Total         por Equipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   630
            TabIndex        =   56
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
         TabIndex        =   51
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
            Picture         =   "OFABRI07.frx":445C
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Consulta de Operaciones Proyectadas por O.F."
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones Proyectadas"
            BeginProperty Font 
               Name            =   "Arial"
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
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Base de Datos de O-F's"
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
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   4955
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   90
         ScaleHeight     =   585
         ScaleWidth      =   4725
         TabIndex        =   19
         Top             =   420
         Width           =   4780
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
            Left            =   4200
            Picture         =   "OFABRI07.frx":45A6
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Anotador de Ordenes de Fabricación"
            Top             =   0
            Width           =   530
         End
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
            Picture         =   "OFABRI07.frx":48B0
            Style           =   1  'Graphical
            TabIndex        =   58
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
            Picture         =   "OFABRI07.frx":4BBA
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Refrescar Ordenes de Fabricación"
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
            Picture         =   "OFABRI07.frx":4D04
            Style           =   1  'Graphical
            TabIndex        =   25
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
            Picture         =   "OFABRI07.frx":500E
            Style           =   1  'Graphical
            TabIndex        =   24
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
            Picture         =   "OFABRI07.frx":5318
            Style           =   1  'Graphical
            TabIndex        =   23
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
            Picture         =   "OFABRI07.frx":5462
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
            Left            =   1050
            Picture         =   "OFABRI07.frx":576C
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
            Left            =   2100
            Picture         =   "OFABRI07.frx":5A76
            Style           =   1  'Graphical
            TabIndex        =   20
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
      BackColor       =   &H00C0E0F8&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   5280
      TabIndex        =   1
      Top             =   210
      Width           =   4425
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   450
         Left            =   840
         ScaleHeight     =   390
         ScaleWidth      =   2820
         TabIndex        =   84
         Top             =   2020
         Width           =   2880
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
            Picture         =   "OFABRI07.frx":5D80
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Ver y Re-Imprimir Copia de Orden de Fabricación Emitida"
            Top             =   -105
            Width           =   530
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta / Re-Impresion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   660
            TabIndex        =   85
            Top             =   105
            Width           =   2160
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0F8&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   315
         TabIndex        =   48
         Top             =   1260
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0F8&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2205
            TabIndex        =   50
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0F8&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   525
            TabIndex        =   49
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
         TabIndex        =   43
         Top             =   420
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
            Left            =   3150
            Picture         =   "OFABRI07.frx":63EA
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Control de Cumplimiento de O/F's"
            Top             =   0
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
            Left            =   1350
            Picture         =   "OFABRI07.frx":66F4
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Pendientes por Sector Productivo"
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
            Picture         =   "OFABRI07.frx":69FE
            Style           =   1  'Graphical
            TabIndex        =   46
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
            Left            =   1950
            Picture         =   "OFABRI07.frx":6D08
            Style           =   1  'Graphical
            TabIndex        =   45
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
            Left            =   2550
            Picture         =   "OFABRI07.frx":6E52
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Consulta / Listado Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B-D de O/F's "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   47
            Top             =   105
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Complementos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   180
      TabIndex        =   0
      Top             =   2940
      Width           =   4955
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1935
         TabIndex        =   40
         Top             =   525
         Width           =   2000
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
            Picture         =   "OFABRI07.frx":715C
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Acceso Directo a Informes de Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes de Producción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   735
            TabIndex        =   42
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
         TabIndex        =   37
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
            Picture         =   "OFABRI07.frx":759E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Acceso Directo a Programación de Carga de Máquinas (MRP-II)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga de Máquinas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   840
            TabIndex        =   39
            Top             =   105
            Width           =   1320
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OFABRI07.frx":78A8
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OFABRI07.frx":7ADC
      TabIndex        =   101
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   99
      ToolTipText     =   "Doble-Click Interrumpe Proceso"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
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
Attribute VB_Name = "OFABRI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INTERRUMPE%
Private Function CALCURESERVAINTE(NORFA$, CI%, DEPODEST%)
        '
        SQLX$ = "SELECT IDSUBOR, CantRes, CANTCONS FROM RESERVA WITH(NOLOCK) WHERE COD_INTE=" & CI% & " AND ((CANTRES-CANTCONS) > 0.005)"
        Set RST = READSET(SQLX$)
        
        With RST
            Do While Not .EOF
                NUORFA$ = SAFETEXT$(!IDSUBOR)
                If ValidaArmador%(NUORFA$) = 0 Then
                    RESEROF = RESEROF + XVALO(!CANTRES)
                    RESEROF = RESEROF - XVALO(!CANTCONS)
                End If
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RST.Close
        Set RST = Nothing
        On Error GoTo 0
        
        If RESEROF < 0 Then RESEROF = 0
        STOKMESA = STODEPOS(CI%, DEPODEST%)
        SOBRANTE = RESEROF - STOKMESA
        If SOBRANTE < 0 Then SOBRANTE = 0
        '
        NECNETA = ROUND#(REDON_INDIVIS(CI%, CDbl(SOBRANTE)), XVALO(CNTDC$))
        '
        CALCURESERVAINTE = NECNETA
        '
End Function

Private Function CALCURESERVAARMA(NORFA$, CI%, DEPODEST%)
        '
        RESEROF = 0
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
        
        SQLX$ = "SELECT IDSUBOR, CantRes, CANTCONS FROM RESERVA WITH(NOLOCK) WHERE COD_INTE=" & CI% & " AND ((CANTRES-CANTCONS) > 0.005)"
        Set RST = READSET(SQLX$)
        
        With RST
            Do While Not .EOF
                NUORFA$ = SAFETEXT$(!IDSUBOR)
                CONDI$ = "IDSUBOR = '" & NUORFA$ & "'"
                ARMADOR% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
                If ARMAD% = ARMADOR% Then
                    RESEROF = RESEROF + XVALO(!CANTRES)
                    RESEROF = RESEROF - XVALO(!CANTCONS)
                End If
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RST.Close
        Set RST = Nothing
        On Error GoTo 0
        
        If RESEROF < 0 Then RESEROF = 0
        STOKMESA = STODEPOS(CI%, DEPODEST%)
        SOBRANTE = RESEROF - STOKMESA
        If SOBRANTE < 0 Then SOBRANTE = 0
        '
        NECNETA = ROUND#(REDON_INDIVIS(CI%, CDbl(SOBRANTE)), XVALO(CNTDC$))
        '
        CALCURESERVAARMA = NECNETA
        '
End Function

Sub CANJE_COMPONENTES()

   '1.-Presentar lista de seleccion de of
       'tamando el deposito configurada.
    Call CARGALISEL("ORFAARMA", "ORDEFABR", "IDSUBOR/A16;COD_EXTE/A18;DESCRIP/A24", "StatuOrf = 1  ORDER BY IDSUBOR ASC", "DISTINCT ")
10  Call SELECHO("ORFAARMA/O-FABRICACION ACTIVAS")
    NORFA$ = VALACT1$("ORFAARMA/Ordenes de Fabricación")
    If NORFA$ = "" Then Exit Sub
    '
    'CARGALISEL CON CODIGO Y LOTE DE LOS COMPONENTES TRANSFERIDOS PARA LA OF
    LOTEORFA$ = ""
    CONDI1$ = "doseccor = '" & NORFA$ & "' AND CODIMOVI = 'TR'  AND CANTINGRE > CANTSALID  AND DEPOMOVI <> " & DEPO%
    If CantRegistros("MOVISTO", CONDI1$) < 1 Then
        Call COMUNI("No Hay Componentes Transferidos Para Esta O-Fabricacion")
        GoTo 10
    End If
    CONDI$ = "doseccor = '" & NORFA$ & "' AND CODIMOVI = 'TR'  AND CANTINGRE > CANTSALID  AND DEPOMOVI <> " & DEPO% & " ORDER BY COD_EXTE ASC "
    Call CARGALISEL("CODLOTOF", "MOVISTO", "COD_EXTE/A18;DESCRIP/A24;IDENLOTE/A16", CONDI$, "DISTINCT ")

20  Call SELECHO("CODLOTOF/Componentes de la Orden " & NORFA$)
    CIXI% = CODINT%(VALACT1$("CODLOTOF"))
    If CIXI% < 1 Then GoTo 10
    LOTEORFA$ = TRIM$(Mid$(VALACT2$("CODLOTOF"), 26))
    '
    Call SELECHO("MASTER/SELECCION DE CODIGO DE REEMPLAZO")
    CI1% = CODINT(VALACT1$("MASTER"))
    If CI1% < 1 Then GoTo 20
    
    'SELECCION DE LOTE PARA EL ARTICULO QUE SE VA A TRANSFERIR
    LOTE$ = ""
    DEPOCEN% = XVALO(CONTROL$("", "DEPOCEN")) ' PRESENTA LOTES DEL DEPOSITO DE PAÑOL
    If ESTRAZABLE%(CI1%) > 0 Then
      Call ARMALILOT(CI1%, DEPOCEN%)
      Call SELECHO("!LILOTSE/Seleccione el Lote de Reemplazo")
      LOTE$ = TRIM$(VALACT1$("!LILOTSE"))
      If LOTE$ = "" Then GoTo 20
    End If

    'INGRESAR CANTIDAD Y DEPOSITO DE ENTRADA
    'PRESENTAR OBJETO PARA QUE INGRESE CANTIDAD A QUITAR Y DEPOSITO DE INGRESO
    '*************************************************************************
    CONDI2$ = "doseccor = '" & NORFA$ & "'  and codimovi = 'TR' AND COD_INTE = " & CIXI% & " AND CantIngre > 0"
    DEPO% = XVALO(VALOBADA("MOVISTO", "DEPOMOVI", CONDI2$))
    Call REPLA(VDEF$, MKI$(CIXI%), 1, 2)
    Call REPLA(VDEF$, LOTE$, 3, 16)
    Call REPLA(VDEF$, Chr$(DEPO%), 23, 1)
    '
12  OBX$ = OBJPLA$("SEDEPCAN/CANTIDAD A REEMPLAZAR", VDEF$)
    If OBX$ = "" Then GoTo 20
    '
    CANTI = CVS(Mid$(OBX$, 19, 4))
    DEPO% = Asc(Mid$(OBX$, 23, 1)) 'DEPOSITO DE SALIDA
    DEPOING% = Asc(Mid$(OBX$, 24, 1)) 'DEPOSITO DE INGRESO
    '
    'VALIDAR DATOS
    If CANTI < 1 Then
      Call COMUNI("Cantidad No Válida")
      VDEF$ = OBX$
      GoTo 12
    End If
    '
    If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
      Call COMUNI("Depósito Salida No Válido")
      VDEF$ = OBX$
      GoTo 12
    End If
   '
   If DEPOING% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
      Call COMUNI("Depósito Ingreso No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   
   '
   'VALIDACION SI NO SUPERA LA CANTIDAD TRANSFERIDA PARA LA ORDEN
   CONDI$ = "doseccor = '" & NORFA$ & "'  and codimovi = 'TR' AND COD_INTE = " & CIXI% & " AND (DEPOMOVI =" & DEPO% & " OR DEPOMOVI=32)"
   If ESTRAZABLE%(CIXI%) > 0 Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTEORFA$ & "'"
   CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
   CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
   CANTTRAN = CANTING - CANTSAL
   '
   If CANTTRAN < CANTI Then
      Call COMUNI("Cantidad Ingresada Supera la Consumida Para Esta Orden ")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   'VALIDACION QUE LA CANTIDAD A TRANSFERIR TENGA STOCKS EN EL DEPOSITO ELEGIDO
   DEPOCEN% = XVALO(CONTROL$("", "DEPOCEN")) ' DEPOSITO CENTRAL
   If ESTRAZABLE(CI1%) > 0 Then 'TRAZABLE
        CANTLOTE = SaldoLote(CI1%, LOTE$, DEPOCEN%)
        If CANTI > CANTLOTE Then
           Call COMUNI("La Cantidad Ingresada Supera el Stock del Lote en Depósito Seleccionado")
           VDEF$ = OBX$
           GoTo 12
        End If
   Else ' NO TRAZABLE
       If STODEPOS(CI1%, DEPOCEN%) < CANTI Then
           Call COMUNI("La Cantidad Ingresada Supera el Stock en Depósito Seleccionado")
           VDEF$ = OBX$
           GoTo 12
       End If
   End If

   '
   'GENERAR MOVIMIENTO DE STOCKS
   '****************************
   CMOV$ = "TR"
   NRO& = ProNroComprobante(CMOV$)
   HOII% = HOY(HOS$)
   FF% = HOII%
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   DOSEC$ = NORFA$
  '
   NUORIT% = 0
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   REFE$ = "Canje de Componentes"
   '
   'MOVIMIENTOS DE SALIDA
   Call MOVISTOK(FF%, CIXI%, LOTEORFA$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, (-1) * CANTI)
   'MOVIMIENTOS DE ENTRADA
   Call MOVISTOK(FF%, CIXI%, LOTEORFA$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOING%, CANTI)
   '
   'MOVIMIENTOS DE SALIDA
   Call MOVISTOK(FF%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOCEN%, (-1) * CANTI)
   'SALIDA DEL DEPOSITO CENTRAL
   Call MOVISTOK(FF%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)

   Screen.MousePointer = 1
   Call COMUNI("Proceso Finalizado")
   
End Sub

Sub CANJE_COMPONENTES_TIPO()
    'RESUMEN:
    'MOSTRAR LISTA DE EQUIPOS
    'MOSTRAR LISTA DE NUMEROS DE SERIES RELACIONADOS
    'MOSTAR LISTA DE TIPO DE  CANJE
    'PASAR INFORMACION A FORMULARIO DE CANJE DE COMPONENTES
    '
10  Call CARGALISEL("EQUIPOS", "MASTER", "CODINT/C16+MASTER/A36", "TRAZABLE = 2 ORDER BY CODIGO ASC")
    Call SELECHO("EQUIPOS/Listado de Equipos")
    CI% = CODINT%(VALACT1$("EQUIPOS"))
    If CI% < 1 Then Exit Sub
    '
20  Call CARGALISEL("LISTNUME", "TANUMSE", "NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", "COD_INTE = " & CI% & " AND (NUMFACTURA = '' OR  NUMFACTURA IS NULL) AND (NUMREMITO = ''  OR NUMREMITO IS NULL) AND DEPONUMSE >= 0  ORDER BY NumeroSerie ASC")
    Call SELECHO("LISTNUME/Números de Serie")
    NUSE$ = VALACT1$("LISTNUME")
    If NUSE$ = "" Then GoTo 10
    '
    Call SELECHO("TATICANJ/Tipos de Canje")
    TIPOCAN% = XVALO(VALACT1$("TATICANJ"))
    If TIPOCAN% < 1 Then GoTo 20
    '
    CANJECOMPO.Label1 = TRIM$(Str$(TIPOCAN%))
    CANJECOMPO.Label7 = TRIM$(NUSE$)
    
    CANJECOMPO.Show 1
    
    GoTo 10
    
    
 
End Sub

'
Sub CONCUMORFAB()
  '27/04/07 (OT-07-0172)
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If PERIO$ = "" Then Exit Sub
  '
  DESDE = FETEXCOR1$(DES%) ' SE MODIFICO EL FORMATO FECHA DE NUMERICO A FECHA CORTA
  HASTA = FETEXCOR1$(HAS%)
  '
  Call ABREORFAB
  '
  SQLX$ = " SELECT FechGen, IDSUBOR, COD_INTE, CanProy, DESTINO, FechEnSol, Fechcierr  FROM ORDEFABR WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE FechGen >= '" & DESDE & "' "
  SQLX$ = SQLX$ + " AND FechGen <= '" & HASTA & "' "
  SQLX$ = SQLX$ + " AND StatuOrf <> 9 "
  SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC"
  '
  'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'  Dim ORDEFABTMP As ADODB.Recordset
'  Set ORDEFABTMP = New ADODB.Recordset
'  ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  '
  Set ORDEFABTMP = READSET(SQLX$)
  With ORDEFABTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
       On Error GoTo 0
       Screen.MousePointer = 1
       Call MENSERR(24, "Ningun Registro Coincide con Registro SQL")
       Exit Sub
    End If
    '
    I& = 0
    Do While Not .EOF
       y$ = REGBLAN("COCUMOF")
       FEMIS% = CVINT(!FechGen)
       Call REPLA(y$, MKI$(FEMIS%), 1, 2)
       Call REPLA(y$, !IDSUBOR, 3, 12)
       Call REPLA(y$, MKI$(!COD_INTE), 15, 2)
       Call REPLA(y$, MKD$(!CanProy), 17, 8)
       Call REPLA(y$, !DESTINO, 25, 8)
       FECHAENT% = CVINT(!FechEnSol)
       Call REPLA(y$, MKI$(FECHAENT%), 33, 2)
       FECHACIE% = CVINT(!Fechcierr)
       Call REPLA(y$, MKI$(FECHACIE%), 35, 2)
       DIATRAZO% = DIENTRE%(FECHAENT%, FECHACIE%)
       Call REPLA(y$, MKI$(DIATRAZO%), 37, 2)
       '
       I& = I& + 1
       Call GRAREG("COCUMOF", y$, I&)
    .MoveNext
    Loop
    '
  End With
  Set ORDEFABTMP = Nothing
  '
  Call SETULTREG("COCUMOF", I&)
  '
  Call HEADERS("COCUMOF", "")
  Call HEADERS("COCUMOF", "Período: " + PERIO$)
  Call FINAL("*COCUMOF")
  '
  '27/04/07 (FIN)
End Sub

Sub GeneraRemiConsig(NORFA$, NPRO%, Optional NUMREMI$)
      '
      Screen.MousePointer = 11

      'GENERA MOVIMIENTO CON CODIGO 0 PARA REGISTRAR EL NUMERO DE REMITO EN CONSIGNACION
      CMO$ = TRIM$(CONTROL$("", "CODRECON"))
      If CMO$ = "" Then CMO$ = "RC"
      '
      UNRO& = ProNroComprobante(CMO$)
      '
      Call BLOQARCH("MOVISTO")
      Call BLOQARCH("TRANSFE")
      '
      Screen.MousePointer = 11
      NC% = (-1) * NPRO%
      VUNI# = 0
      MONE$ = ""
      LOTE$ = ""
      FECHA% = HOY(HOS$)
      DOPRI$ = "RC-" + TRIM$(Str$(UNRO&))
      DOSEC$ = "RC-" + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
      While Len(DOSEC$) < 10
          DOSEC$ = DOSEC$ + " "
      Wend
      NORI% = 1
      CI1% = 0
      CAN = 0
      REFE$ = "RC.a " + "Armador Externo"
      Call MOVISTOK(FECHA%, CI1%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
      'FIN DE MOVIMIENTO CON CODIGO 0
      '
      REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO&))
      HOII% = HOY(HOS$)
      FF% = HOII%
      '
      'IMPRESION DE REMITO DE CONSIGNACION
      '
      If NPRO% <> 0 Then GoTo 22
      CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
      ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
      TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
      NPRO% = CVI(Mid$(TX$, 68, 2))
      NUOCO$ = 0

22    FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
      If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Screen.MousePointer = 1

                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                Exit Sub
            End If
      End If
      '
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = TRIM$(REMI$)
      TIPODOC$ = "Remito de Traslado"
      DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
      DATRANSPO$ = DTP$

      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
      CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
      CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
      CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
      CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
      CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
      ' IMPRESION DEL CODIGO DE BARRAS PARA PICKING DEL REMITO EN CONSIGNACION
        NUCL$ = Right("000000" & TRIM$(Str(NPRO%)), 6)
        NUREMI$ = Right("00000000" & TRIM$(Mid$(REMI$, 4)), 8) 'MODIFIC. P'LONGITUD FIJA
        CADENA$ = NUCL$ + NUREMI$
        BARRA$ = CODIBARRA$(CADENA$, 2)
      CODPARAM$(9) = "REF-MAT3": DOCPARAM$(9) = BARRA$
      '
      CAPARDOC% = 9
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "CAKILOS"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "UBI-MPRI"
      CODTABU1$(8) = "NRO-LOTE"
      CODTABU1$(9) = "REF-MAT6"
      
      CACOLTAB1% = 9
      '
      DESTIDOC% = NPRO% * (-1)
      CAITAB1% = 0
      '
      FORVALE$ = "VALEMAT"
      If InStr(1, EMPREAC$, "GABAL") > 0 Then FORVALE$ = "VALEMATG"
      For U& = 1 To ULTREG&("!" & FORVALE$)
         XX$ = REGLEIDO$("!" & FORVALE$, U&)
         If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
            '
            CANS$ = Str$(CVS(Mid$(XX$, 19, 4)))
            CICRU% = CVI(Mid$(XX$, 1, 2))
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            UNIMEX$ = UNIMED$(CICRU%)
            LOTE$ = TRIM$(Mid$(XX$, 49, 16))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT$(CICRU%)
            TETABU1$(3, CAITAB1%) = CANS$
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = UNIMEX$
            TETABU1$(6, CAITAB1%) = NUOCO$
            TETABU1$(7, CAITAB1%) = UBISTO$(CICRU%)
            TETABU1$(8, CAITAB1%) = LOTE$
            TETABU1$(9, CAITAB1%) = CANBULTOS(CICRU%, XVALO(CANS$), "")
            '
         End If
      Next U&
      '
      NUMREMI$ = DOPRI$
      '
      Screen.MousePointer = 1
      '
      FORIPRE$ = "FOREMICO"
      If TRIM$(CONTROL$("", "RCONSIOF")) <> "" Then FORIPRE$ = TRIM$(CONTROL$("", "RCONSIOF"))
      Call PRINTDOC("@" + FORIPRE$) ' Remito en Consignacion
      '
End Sub

Sub GRABAETIQUETASGABAL(NORFA$, MODO%)
        '
        CONDI$ = "IdSubOr='" & NORFA$ & "'"
        CAN0# = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$))
        CIOF0% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$))
        If MODO% = 1 Then
            CANTI_ETIQ_A_GENERAR% = CANBULTOS(CIOF0%, CAN0#, "VENTA")
            SQLX$ = "SELECT * FROM ORFAETI WHERE IDSUBOR= ''"
            Dim ORDEFABTMP As ADODB.Recordset
            Set ORDEFABTMP = New ADODB.Recordset
            ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            '
            With ORDEFABTMP
                For I% = 1 To CANTI_ETIQ_A_GENERAR%
                    '
                    .AddNew
                    !IDSUBOR = TRIM$(NORFA$)
                    !NUORDEN = I%
                    !Status = 0
                    !Usuario = USERNAME$
                    !MARBORRA = 0
                    !CodiEmpr = CodiEmp%
                    .Update
                    '
                Next I%
            End With
            On Error Resume Next
            ORDEFABTMP.Close
            Set ORDEFABTMP = Nothing
            On Error GoTo 0
            '
        ElseIf MODO% = 2 Then
20             RNDX1$ = RANDSTRING$(-10)
               CONDI$ = "IDSUBOR = '" & RNDX1$ & "'"
               If CantRegistros("ORFAETI", CONDI$, "NOMESS") > 0 Then GoTo 20
               Call CreaRegistro("ORFAETI", "IDSUBOR ", RNDX1$, "NOMESS")
                    ORD% = XVALO(VALOBADA("ORFAETI", "MAX(NUORDEN)", "IDSUBOR='" & NORFA$ & "'")) + 1
               Call ACTUREGISTRO("ORFAETI", "NUORDEN", CONDI$, ORD%, REGAF&, "NOMESS")
               Call ACTUREGISTRO("ORFAETI", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
               Call ACTUREGISTRO("ORFAETI", "USUARIO", CONDI$, USERNAME$, REGAF&, "NOMESS")
               Call ACTUREGISTRO("ORFAETI", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
               '
               Call ACTUREGISTRO("ORFAETI", "IDSUBOR", CONDI$, NORFA$, REGAF&, "NOMESS")
               '
               XANTI = LINDIVEN(CIOF%)
               CANTI = CAN0# + XANTI
               Call ACTUREGISTRO("ORDEFABR", "CANPROY", "IDSUBOR='" & NORFA$ & "'", CANTI, REGAF&)
        End If
End Sub

Sub REEMPLAZO_COMPONENTES()
    '
    'AGREGAR PADR_FUN_SQL ,STDSCREEN, FORM_MGR_SQL,
    
    Call CARGALISEL("INDINUSE", "TANUMSE", "NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", "NUMEROSERIE <> '' ORDER BY FECHALTA ASC, NUMEORD ASC")
10  Call SELECHO("INDINUSE")
    NUSE$ = VALACT1$("INDINUSE")
    If NUSE$ = "" Then Exit Sub
    '
    FILTX$ = TRIM$(NUSE$)
    Call STDFORM("REEMCO ", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 10
    End If
    '
    NORFA$ = VALOFORM("PARAM-4")
    CIXI% = XVALO(VALOFORM("PARAM-2"))
    
20  OPX% = COMALTER%("Opciones de Trabajo \\Quitar Componentes\Agregar Componentes")
    If OPX% < 1 Or OPX% > 2 Then GoTo 10
    If OPX% = 1 Then
      Call QUITACOMP(NORFA$)
      GoTo 20
    ElseIf OPX% = 2 Then
      Call AGREGACOMP(NORFA$)
      GoTo 20
    End If


    
End Sub

Sub AGREGACOMP(NORFA$)
   '
   'SELECCCIONA CODIGO Y LOTE EN CASO QUE SEA TRAZABLE
   DEPOCEN% = XVALO(CONTROL$("", "DEPOCEN"))
   SELCODLOT07.Label15 = DEPOCEN%
10 SELCODLOT07.Show 1
    '
    'VALIDO SI COMPLETÓ DATOS
   CIXI% = CODINT%(TRIM$(SELCODLOT07.Text1))
   If CIXI% < 1 Then Exit Sub
    '
   LOTE$ = TRIM$(SELCODLOT07.Label28)
   
   'PRESENTAR OBJETO PARA QUE INGRESE CANTIDAD A QUITAR Y DEPOSITO DE INGRESO
    '*************************************************************************
   Call REPLA(VDEF$, MKI$(CIXI%), 1, 2)
   Call REPLA(VDEF$, LOTE$, 3, 16)

12 OBX$ = OBJPLA$("CANTIDEP/CANTIDAD A AGREGAR", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   CANTI = CVS(Mid$(OBX$, 19, 4))
   DEPO% = Asc(Mid$(OBX$, 23, 1))
   '
   'VALIDAR DATOS
   If CANTI < 1 Then
      Call COMUNI("Cantidad No Válida")
      VDEF$ = OBX$
      GoTo 12
   End If
   If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
      Call COMUNI("Depósito No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   'VALIDAR QUE EL LOTE TENGA CANTIDAD SUFICIENTE PARA AGREGAR
   If ESTRAZABLE(CIXI%) > 0 Then 'TRAZABLE
        CANTLOTE = SaldoLote(CIXI%, LOTE$, DEPO%)
        If CANTI > CANTLOTE Then
           Call COMUNI("La Cantidad Ingresada Supera el Stock del Lote en Depósito Seleccionado")
           VDEF$ = OBX$
           GoTo 12
        End If
   Else ' NO TRAZABLE
       If STODEPOS(CIXI%, DEPO%) < CANTI Then
           Call COMUNI("La Cantidad Ingresada Supera el Stock en Depósito Seleccionado")
           VDEF$ = OBX$
           GoTo 12
       End If
   End If
   '
   'GENERAR MOVIMIENTO DE STOCKS DE SALIDA
   '**************************************
   CMOV$ = "CF"
   HOII% = HOY(HOS$)
   FF% = HOII%
   DOPRI$ = NORFA$
   DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
  '
   NUORIT% = 0
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   REFE$ = "Agregado de Componentes"
   '
   Call MOVISTOK(FF%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI * (-1))
   '
   Call COMUNI("Proceso Finalizado")
   Screen.MousePointer = 1

End Sub


Sub QUITACOMP(NORFA$)
    'QUITAR COMPONENTES
    '1.- PRESENTAR CARGALISEL CON CODIGO Y DESCRIPCION Y LOTE DE LOS COMPONENTES CONSUMIDOS
    'PARA LA OF ELEGIDA
    '2.-PRESENTAR OBJETO PARA QUE INGRESE CANTIDAD A QUITAR Y DEPOSITO DE INGRESO
    '3.-GENERAR MOVIMIENTO DE ENTRADA
    '***********************************************************************************
    
    'PRESENTAR CARGALISEL CON CODIGO Y DESCRIPCION Y LOTE DE LOS COMPONENTES CONSUMIDOS
    '**********************************************************************************
    CONDI$ = "dopricor = '" & NORFA$ & "' " & " and (codimovi = 'CF' or codimovi = 'TR')ORDER BY cod_exte,idenlote "
    Call CARGALISEL("CODLOTE", "MOVISTO", "COD_EXTE/A20;DESCRIP/A30;IDENLOTE/A16", CONDI$, "DISTINCT")
    
10  Call SELECHO("CODLOTE/SELECCIÓN DE CODIGO Y LOTE PARA " & NORFA$)
   
    CODCOMP% = CODINT(VALACT1$("CODLOTE"))
    DESCRIP$ = TRIM$(Mid$(VALACT2$("CODLOTE"), 1, 30))
    LOTE$ = TRIM$(Mid$(VALACT2$("CODLOTE"), 32, 16))
    If CODCOMP% < 1 Then: Exit Sub
    '
    'PRESENTAR OBJETO PARA QUE INGRESE CANTIDAD A QUITAR Y DEPOSITO DE INGRESO
    '*************************************************************************
    Call REPLA(VDEF$, MKI$(CODCOMP%), 1, 2)
    Call REPLA(VDEF$, LOTE$, 3, 16)

12 OBX$ = OBJPLA$("CANTIDEP/CANTIDAD DEVUELTA", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   CANTI = CVS(Mid$(OBX$, 19, 4))
   DEPO% = Asc(Mid$(OBX$, 23, 1))
   '
   'VALIDAR DATOS
   If CANTI < 1 Then
      Call COMUNI("Cantidad No Válida")
      VDEF$ = OBX$
      GoTo 12
   End If
   If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
      Call COMUNI("Depósito No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   'VALIDAR QUE LA CANTIDAD NO SEA MAYOR A LA CONSUMIDA PARA ESTA ORDEN
   CONDI$ = "COD_INTE = " & CODCOMP% & "  AND dopricor ='" & NORFA$ & "' " & " and (codimovi = 'CF' or codimovi = 'TR') "
   If LOTE$ <> "" Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
   '
   CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
   CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
   CANTCONSU = CANTSAL - CANTING
   '
   If CANTI > CANTCONSU Then
      Call COMUNI("Imposible Quitar Componentes \La Cantidad Consumida es Menor a la Cantidad Ingresada")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   'GENERAR MOVIMIENTO DE STOCKS
   '****************************
   CMOV$ = "CF"
   HOII% = HOY(HOS$)
   FF% = HOII%
   DOPRI$ = NORFA$
   DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
  '
   NUORIT% = 0
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   REFE$ = "Quita de Componentes"
   '
   Call MOVISTOK(FF%, CODCOMP%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
   '
   Call COMUNI("Proceso Finalizado")
   Screen.MousePointer = 1

End Sub

Sub TRANSF_A_ARMADOR(NUORFAB$, ARMAD%)
   Dim MATRANSFE#(32767)
   Dim CANECE(32767)
   '
   Call ABRESTOCKS
   ' VALIDACIONES
   If TRIM$(NUORFAB$) = "" Then Exit Sub
   'EN BASE AL ARMADOR BUSCO EL PROVEEDOR Y VALIDO
If ValidaArmador%(NUORFAB$) > 0 Then
   TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
   ARMADOR$ = TRIM$(Mid$(TX$, 3, 64))
   proveed% = CVI(Mid$(TX$, 68, 2))
   TIPOARM$ = TRIM$(UCase$(Mid$(TX$, 67, 1)))
   Deposito% = Asc(Mid$(TX$, 71, 1)) ' GABAL - LEE EL DEPOSITO ASIGNADO AL ARMADOR
   '
   If proveed% < 1 Then
       Call COMUNI("Falta Asignar Proveedor al Armador " & ARMAD% & "\" & ARMADOR$)
       Exit Sub
   End If
   '
   DEPODEST% = XVALO(CONTROL("", "DEPOARMA"))
   If DEPODEST% < 1 Then
      Call COMUNI("Debe Configurar un Depósito de Armadores Externos")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   'PONE EN 0 LOS CAMPOS CANTINGRE Y CANTSALID PARA EVITAR VALORES NULL
   Call PoneEnCeroNull("MOVISTO", "CANTINGRE", REGAF&, "NOMESS")
   Call PoneEnCeroNull("MOVISTO", "CANTSALID", REGAF&, "NOMESS")

   'VALIDACION QUE NO HAYA MATERIAL TRANSFERIDO AL DEPOSITO TALLER INTERNO
   DepOrig% = CONTROL("", "DEPODEST")         ' TALLER INTERNO
   DEPODEST% = XVALO(CONTROL("", "DEPOARMA")) ' DEPOSITO DE ARMADORES
   '
    'comentado para que permita cambiar de deposito en el caso que ya tenga material transferido

''   For t& = 1 To NUMAS%: MATRANSFE#(t&) = 0: Next t&
''   '
''   'VALIDA QUE NO HAYAN REALIZADO TRANSFERENCIAS PREVIAS AL DEPOSITO DE TALLER INTERNO
''   Call SETULTREG("!VALEMAT", 0)
''   OPX% = 0
''   SQLX$ = "SELECT SUM(CANTINGRE - CANTSALID)AS MaterTrans, COD_INTE, IdenLote "
''   SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
''   SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NUORFAB$ & "' "
''   SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
''   SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
''   SQLX$ = SQLX$ + " GROUP BY COD_INTE, IdenLote "
''   Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
''   '
''   JJ& = 0
''   With MoviTemp
''      Do While Not .EOF
''        CI% = XVALO(!Cod_Inte)
''        IDLO$ = SAFETEXT$(!idenLote)
''        MATRAN# = XVALO(!MaterTrans)
''        If MATRAN# > FRACMIN(CI%) Then
''           ZZ$ = REGBLAN$("VALEMAT")
''           Call REPLA(ZZ$, MKI$(CI%), 1, 2)
''             CANTID = MATRAN#
''           Call REPLA(ZZ$, MKS$(CANTID), 19, 4)
''           Call REPLA(ZZ$, "......", 27, 6)
''           Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
''           Call REPLA(ZZ$, IDLO$, 49, 16)
''           JJ& = JJ& + 1
''           Call GRAREG("!VALEMAT", ZZ$, JJ&)
''        End If
''      .MoveNext
''      Loop
''   End With
''   MoviTemp.Close
''   Set MoviTemp = Nothing
''   '
''   Screen.MousePointer = 1
''
''   If JJ& > 0 Then
''      OPX% = COMALTER%("Atencion: Esta Orden de Fabricación (Asignada a Armador Externo)\ya Tiene Material Transferido al Taller Interno: " & DepOrig% & "\  \Este Material Debe Transferirse al Armador Externo\\Cancelar esta Transferencia\Transferir al Armador Externo")
''      If OPX% <> 2 Then ' SI NO TRANSFIERE
''          Exit Sub
''      End If
''      GoTo 50
''   End If
   '
   'Transferencia normal DE DEPOSITO CENTRAL a Deposito de Armadores
   DepOrig% = XVALO(CONTROL$("", "DEPOCEN"))
   DEPOFAB% = XVALO(CONTROL("", "DEPODEST"))
   If InStr(1, EMPREAC$, "GABAL") > 0 Then DEPODEST% = Deposito% ' GABAL - TOMA EL DEPOSITO ASIGNADO AL ARMADOR
   '
   VDEF$ = String$(64, 0)
   Call REPLA(VDEF$, NUORFAB$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPODEST%), 14, 1)
   Call REPLA(VDEF$, Chr$(DepOrig%), 13, 1)
   '
   DEPARMA% = DEPODEST% ' PARA DESPUES COMPARAR QUE NO CAMBIE DEP0SITO

12 OBJ$ = "SEORFADE"
   If InStr(1, EMPREAC$, "GABAL") > 0 Then OBJ$ = "SEORFADG"
   OBX$ = OBJPLA$(OBJ$, VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   '
   If DEPODEST% <> DEPOFAB% Then
     Screen.MousePointer = 11
     MATERIAL_TRANS_54% = 0
     SQLX$ = "SELECT SUM(CANTINGRE - CANTSALID)AS MaterTrans, COD_INTE, IdenLote "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NUORFAB$ & "' "
       SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
       SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
       SQLX$ = SQLX$ + " GROUP BY COD_INTE, IdenLote "
       Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       JJ& = 0
       With MoviTemp
           Do While Not .EOF
             MATRAN# = XVALO(!MaterTrans)
             If MATRAN# > FRACMIN(CI%) Then
                MATERIAL_TRANS_54% = 1
                Exit Do
             End If
           .MoveNext
           Loop
       End With
       MoviTemp.Close
       Set MoviTemp = Nothing
       '
       Screen.MousePointer = 1
       '
       If MATERIAL_TRANS_54% > 0 Then
         If COMALTER%("Verificar Depósito a Transferir Materiales\Orden de Fabricación Posee Material Transferidos a Depósito: " & DEPOFAB% & " \\Modificar\Confirmar Transferencia") <> 2 Then GoTo 12
       End If
   End If
    
   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DepOrig% <> DEPOCENTRAL% Then
      If DERACCE%("CAMBIDEP", "Cambiar Depositos de Transferencia") < 2 Then GoTo 12
   End If
   '
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then
      Call MENSERR(24, "Depósito Origen No Válido"): GoTo 12
   End If
   '
   'comentado para que permita cambiar de deposito en el caso que ya tenga material transferido
'   'VALIDA QUE NO CAMBIE EL DEPOSITO DEL ARMADOR .
'   If DEPODEST% <> DEPARMA% Then
'      Call MENSERR(24, "No Válido Cambiar de Depósito\Orden de Fabricación AsignadA a Armador Externo\Depósito de Consumo: " & DEPARMA%)
'      GoTo 12
'   End If

   'RECORRE LA TABLA RESERVAS Y ARMA EL ARCHIVO RESERVAS
   CONDI$ = "CantRes - CantCons > 0 "
   CONDI$ = CONDI$ + " AND IDSUBOR = '" & NUORFAB$ & "'"
   '
   FIN& = CantRegistros("RESERVA", CONDI$)
   If FIN& < 1 Then
       Call MENSERR(24, " No Existen Reservas Efectuadas.\para esta Orden de Fabricacion.")
       GoTo 12
   End If
   Screen.MousePointer = 11

   SQLX$ = " SELECT COD_INTE, IDSUBOR, FechRes, CantRes, CANTCONS, Saldo_Entre, IdenLote, Deposito, UsoUnit FROM RESERVA WITH(NOLOCK) WHERE " + CONDI$
'   Dim RESERTMP As ADODB.Recordset
'   Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set RESERTMP = READSET(SQLX$)
   JJ& = 0: CONTA% = 0
   REBLA$ = REGBLAN("!RESERVA")
   Call SETULTREG("!RESERVA", 0)
   With RESERTMP
      Do While (.EOF = False)
          y$ = REBLA$
          CII% = XVALO(!COD_INTE)
          If ESTRAZABLE(CII%) = 2 Then
              CONTA% = CONTA% + 1 ' valida si hay codigos trazables
            Else
              Call REPLA(y$, MKI$(!COD_INTE), 1, 2)
              Call REPLA(y$, !IDSUBOR, 3, 12)
              Call REPLA(y$, MKI$(!COD_INTE), 15, 2)
              Call REPLA(y$, MKI$(CVINT(!FechRes)), 17, 2)
              Call REPLA(y$, MKD$(!CANTRES), 19, 8)
              Call REPLA(y$, MKD$(!CANTCONS), 27, 8)
              Call REPLA(y$, MKD$(!Saldo_Entre), 35, 8)
              Call REPLA(y$, SAFETEXT$(!IDENLOTE), 43, 16)
              Call REPLA(y$, Chr$(!Deposito), 60, 1)
              Call REPLA(y$, MKS$(!USOUNIT), 75, 4)
              '
              JJ& = JJ& + 1
              Call GRAREG("!RESERVA", y$, JJ&)
          End If
13    .MoveNext
      Loop
   End With
   RESERTMP.Close
   Set RESERTMP = Nothing
   Call SETULTREG("!RESERVA", JJ&)
   Call CIERRARCH("!RESERVA")
   '
   Screen.MousePointer = 1

   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\los cuales Seran Excluidos de la Lista de Componentes")
   End If
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("!RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!RESERVA", U&)
      ORFA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) '- CVD(Mid$(XX$, 27, 8))
      '
      If CANTI < 0 Then CANTI = 0
      If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then CANTI = CANTI / 1000
      '
      IDLO$ = TRIM$(Mid$(XX$, 43, 16))
      CONDI$ = " DoSecLar = '" & TRIM$(ORFA$) & "'"
      CONDI$ = CONDI$ & " AND Cod_Inte = " & CI%
      'CONDI$ = CONDI$ & " AND DepoMovi = " & DEPODEST%
      'SE MODIFICA LA VERIFICACION DE LO TRANSFERIDO TANTO EN EL DEPOSITO DE ARMADORES (32)
      'COMO EN EL DE DESTINO DE FABRICA (54). ESTO PORQUE AL TRANSFERIR POR PRIMERA VEZ
      'EL MATERIAL ESTA EN EL DEPOSITO DE ARMADORES, PERO AL RECIBIR LOS EQUIPOS EL MODULO DE RECEPCION
      'PASA LOS COMPONENTES AL DEPOSITO DE DESTINO DE FABRICA (54)
      '
      'SE CAMBIA PARA QUE CALCULE SOBRE EL
      'DEPOSITO CONFIGURADO, MAS ALLA QUE CAMBIE EN EL OBJPLA POR EL 54 U OTRO
      '
      CONDI$ = CONDI$ & " AND (DepoMovi = " & DEPARMA% & " OR DepoMovi = " & DEPOFAB% & ")"
      CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "'"
      CANTXX = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
      CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
      STODEST = CANTXX - CANTSAL 'STODEPOS(CI%, DEPODEST%)
      ' STOCENT = STODEPOS(CI%, DepOrig%)
      STOCENT = (SaldoLote(CI%, IDLO$, DepOrig%))
      If STOCENT < 0 Then STOCENT = 0
      NECNETA = CANTI - STODEST
      If NECNETA < 0 Then NECNETA = 0
      
      If InStr(1, EMPREAC$, "GABAL") > 0 Then
        'CALCULO DEL SOBRANTE
        RESERVAS = CALCURESERVAARMA(NUORFAB$, CI%, DEPODEST%)
        STODEST = STODEPOS(CI%, DEPODEST%)
        NECNETA = RESERVAS
      End If
      
      VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
      If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
         VALEMATX = STOCENT
      End If
      '
      If VALEMATX < 0.05 Then VALEMATX = 0
       '
      ZZ$ = REGBLAN$("VALEMAT")
      Call REPLA(ZZ$, MKI$(CI%), 1, 2)
      Call REPLA(ZZ$, MKS$(CANECE(CI%)), 3, 4) ' NECESIDAD BRUTA
        If InStr(1, EMPREAC$, "GABAL") > 0 Then Call REPLA(ZZ$, MKS$(CANTI), 3, 4)
      Call REPLA(ZZ$, MKS$(STODEST), 7, 4)     ' STOCK DEPOSITO DESTINO
      Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
      Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
      Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
      Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
      Call REPLA(ZZ$, "......", 27, 6)
      Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
      Call REPLA(ZZ$, IDLO$, 49, 16)
      Call GRAREG("!VALEMAT", ZZ$, U&)
   Next U&
   Call SETULTREG("!VALEMAT", FIN&)
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1

25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", Chr$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
   VTB% = VENTABU%("VALEMAT9" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then Exit Sub
   '
   Screen.MousePointer = 11

   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CONTA% = 0
   For JJ1& = 1 To ULTREG&("!VALEMAT") ' valida si hay codigos trazables
      XX$ = REGLEIDO$("!VALEMAT", JJ1&)
      CI% = CVI(Mid$(YY$, 1, 2))
      If ESTRAZABLE(CI%) = 2 Then
        CONTA% = CONTA% + 1
      End If
   Next JJ1&
   '
   Screen.MousePointer = 1

   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Deben ser Excluidos de la Lista de Componentes")
      GoTo 25
   End If
   '
   Screen.MousePointer = 11

   If CONTROL("VANEDECE", "DEPOCEN") = "SI" Then ' SI CONFIGURACION DE CONTROL DE NEGATIVOS
      For YX& = 1 To ULTREG&("!VALEMAT")
         YY$ = REGLEIDO$("!VALEMAT", YX&)
         CI% = CVI(Mid$(YY$, 1, 2))
         CAN# = CVS(Mid$(YY$, 19, 4))
             If CAN# > 0.005 Then
               STOCXCENT# = STODEPOS(CI%, DepOrig%) 'CVS(Mid$(YY$, 15, 4))
               VALMAT# = CVS(Mid$(YY$, 19, 4))
               If STOCXCENT# < VALMAT# Or STOCXCENT# < 0 Then
                 Call MENSERR(24, "Transferencia No Procesable\Código: " & CODEXT$(CI%) & "\Generará Un Stock Negativo en Depósito de Origen")
                 GoTo 25
               End If
             End If
      Next YX&
   End If
   
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then Exit Sub
   ' GRABAR MOVIMIENTO DE TRANSFERENCIA
   '
50 CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   Screen.MousePointer = 11

   If CMOV$ = "" Then CMOV$ = "TR"
   '
   Call BLOQARCH("MOVISTO")
   NRO& = ProNroComprobante(CMOV$)
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   DOSEC$ = TRIM$(NUORFAB$) ' NUMERO DE ORDEN DE FABRICACION
   While Len(DOSEC$) < 10:      DOSEC$ = DOSEC$ + " ":   Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         NORI% = NORI% + 1
         '
         LOTE$ = TRIM$(Mid$(XX$, 49, 16))
         '
         REFE$ = "TR.a " + Left$(ARMADOR$, 20)
         CANTI = (-1) * CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         '
      End If
   Next U&
   '
   Call BAJALDISCO
   '
   '
   'IMPRIME EL VALE DE TRANSFERENCIA
   '
   CANCELPRINT% = 0
   IMPR$ = TRIM$(CONTROL$(IDENTER$, "SUIMPVAL"))
   If IMPR$ = "SI" Then CANCELPRINT% = 1

60 FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(Str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
      CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUORFAB$)
      CAPARDOC% = 3
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "NRO-LOTE"
      CACOLTAB1% = 7
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
            LOTE$ = TRIM$(Mid$(XX$, 49, 16))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = UNIMED$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
            TETABU1$(7, CAITAB1%) = LOTE$
            '
         End If
      Next U&
      '
      If CAITAB1% < 1 Then
         Screen.MousePointer = 1
         Exit Sub
      End If
      '
      Call PRINTDOC("FORVALMA")   ' Vale de Materiales
      Call BAJALDISCO
      '
   End If
   '
   Screen.MousePointer = 1

   Call GeneraRemiConsig(NUORFAB$, proveed%, DOPRI$)
   '
   Call COMUNI("Transferencia Realizada a Depósito de Armador")
   End If
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub AnulOrd_Click()
Call Command45_Click
End Sub

Private Sub AnuOrd_Click()
Call Command45_Click
End Sub

Private Sub CanceReser_Click()
Call RESERV07.Command4_Click
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
RESERV07.Option1 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
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
           j& = 0
           Screen.MousePointer = 11
           '
           For U& = 1 To ULTREG&("ORDEFABR")
               x$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(x$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(x$, 89, 8))
                     CAPRA# = CVD(Mid$(x$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     If CAPEN# > 0 Then
10                      SPX$ = Mid$(x$, 113, 4)
                        If ECOARCH$("SECPROD", SPX$) = "" Then
                          CIXI% = CVI(Mid$(x$, 83, 2))
                          OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + String$(4, 0))
                          If OBX$ = "" Then GoTo 99
                          SPX$ = Mid$(OBX$, 3, 4)
                          '///***///
                          'edicion de Sector de O/F's
                          OF1$ = Mid$(x$, 69, 12)
                          Call ABREORFAB
                          'With OrdeFabr
                            'SQLX$ = " IdSubOr = '" & OF1$ & "'"
                            '.FindFirst SQLX$
                            SQLX$ = "UPDATE OrdeFabr SET "
                            SQLX$ = SQLX$ + " SecciPro = '" & CVS(SPX$) & "'"
                            SQLX$ = SQLX$ + " WHERE IdSubOr = '" & OF1$ & "'"
                             FLEXCONN.Execute (FILTSQL$(SQLX$))
'14                          If .NoMatch = False Then
'                              .Edit
'                              !SecciPro = CVS(SPX$)
'                              .Update
'                              .FindNext SQLX$
'                              GoTo 14
'                            End If
'                          End With
                          '///***///
                          Call REPLA(x$, SPX$, 113, 4)
                          Call GRAREG("ORDEFABR", x$, U&)
                          GoTo 10
                        End If
15                      If SPX$ = SESEL$ Then
                           Call REPLA(x$, MKD$(CAPEN#), 159, 8)
                           Call REPLA(x$, MKD$(0), 167, 8)
                           j& = j& + 1
                           Call GRAREG("ORFASEL", x$, j&)
                        End If
                     End If
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", j&)
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
               REFEX8$ = TRIM$(Mid$(XX$, 43, 30))
             Call REPLA(YY$, REFEX8$, 23, 30)
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
    Command14.Enabled = False
    If DERACCE%("STPOFAB", "Configuracion de Funcionamiento de O/Fabricacion") = 2 Then
        OPORFA07.Show 1
    End If
    Command14.Enabled = True
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

Private Sub COMMAND3_Click()
   COMMAND3.Enabled = False
   ACONEC$ = "REPARA02"
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "ARCON") > 0 Then ACONEC$ = "REPABO03"
   Call CONECRUN(ACONEC$, "Reparaciones")
   COMMAND3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call REFRESHECO
    Call REPRINORFA
    Command4.Enabled = True
End Sub

Sub REFRESHECO()
    '
    Dim RST As ADODB.Recordset
    Set RST = READSET("SELECT IDSUBOR, CanProy, COD_INTE, StatuOrf FROM ORDEFABR WITH(NOLOCK) ORDER BY IDSUBOR")
    FIN& = CantRegistros("ORDEFABR", "IDSUBOR <> ''", "NOMESS")
    '
    REGF& = 0: REGP& = 0
    With RST
        Do While Not .EOF
            Call TRACE(24, REGF&, FIN&)
            '
            NORFA$ = SAFETEXT$(!IDSUBOR)
            CANTI# = XVALO(!CanProy)
            CIXI% = XVALO(!COD_INTE)
            STAT% = XVALO(!StatuOrf)
            DESCRI$ = TRIM$(CODEXT$(CIXI%)) + " - " + TRIM$(DESCRIT$(CIXI%))
            CANS$ = " - (" + TRIM$(Str$(CANTI#)) + ")"
            DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
            ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
            '
            XX$ = REGBLAN$("ECORDEF")
            Call REPLA(XX$, NORFA$, 1, 12)
            Call REPLA(XX$, DESCRI1$, 13, 52)
            '
            REGF& = REGF& + 1
            Call GRAREG("ECORDEF", XX$, REGF&)
            If STAT% = 1 Then
                REGP& = REGP& + 1
                Call GRAREG("ECORDEP", XX$, REGP&)
            End If
            
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    '
    Call SETULTREG("ECORDEF", REGF&)
    Call SETULTREG("ECORDEP", REGP&)
    '
    Call CIERRARCH("ECORDEF")
    Call CIERRARCH("ECORDEP")
    '
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
    Command41.Enabled = False
    Call ORFNUE(1)
    Command41.Enabled = True
End Sub

Private Sub Command49_Click()
   ANOTAOR07.Show 1
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If InStr(1, EMPREAC$, "GABAL") > 0 Then
5       OPX% = COMALTER%("Opciones de Impresión\\Orden de Fabricación\Vale de Materiales de Armadores")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        If OPX% = 1 Then GoTo 20
        '
10      Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
        NORFA$ = TRIM$(VALACT1$("ECORDEF"))
        If NORFA$ = "" Then GoTo 5
        NPX& = XVALO(Mid$(NORFA$, 4, 6))
        CONDI$ = "IDSUBOR='" & NORFA$ & "'"
        CI% = XVALO(VALOBADA("ORDEFABR", "Cod_inte", CONDI$))
        NUOCONUE& = XVALO(VALOBADA("ORDEFABR", "Nume_Ocom", CONDI$))
        CANTIAFAB = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
        CAIT% = CantRegistros&("FORMULAS", "CodIConj =" & CI%)
        If CAIT% > 0 Then
           FILTX$ = TRIM$(Str$(CI%)) & ";" & TRIM$(Str$(NUOCONUE&)) & ";" & TRIM$(Str$(CANTIAFAB))
           Call STDFORM("LISTAMAT", FILTX$, "PRINT")
        Else
            Call COMUNI("El Producto Fabricado No Posee Fórmula de Composición")
        End If
        GoTo 10
        '
20      Call VERORFA
        GoTo 5
        '
   End If
   '
50 Call VERORFA
99 Command5.Enabled = True
End Sub

Private Sub Command50_Click()
   Call ABREORFAB
   SQLX$ = " SELECT * FROM ORDEFABR WHERE SecciPro < 0 "
   '
   Dim ORDEFABTMP As ADODB.Recordset
   Set ORDEFABTMP = New ADODB.Recordset
   ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   Screen.MousePointer = 11
   With ORDEFABTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.\Imposible continuar la carga")
        Exit Sub
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     '
     Do While Not .EOF
        J3& = J3& + 1
        Call TRACE(20, J3&, FIN&)
        If XVALO(!DESTINO) >= 1 Then
           SEPRO = XVALO(!DESTINO)
           !SecciPro = SEPRO
           !DESTINO = "STOCK"
           .Update
       End If
     .MoveNext
     Loop
   End With
   '
   ORDEFABTMP.Close
   Set ORDEFABTMP = Nothing
   Screen.MousePointer = 1
   Call COMUNI("Completado")
   '
End Sub

Private Sub Command51_Click()
    Command51.Enabled = False
    ACONEC$ = "IFABRI07"
    Call CONECRUN(ACONEC$, "Informes de Producción")
    Command51.Enabled = True
End Sub

Private Sub Command52_Click()
    Command52.Enabled = False
    ACONEC$ = "CARMAQ07"
    Call CONECRUN(ACONEC$, "Carga de Máquinas")
    Command52.Enabled = True
End Sub




Private Sub Command53_Click()
  Command53.Enabled = False
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If PERIO$ = "" Then Exit Sub
  '
  DESDE = FETEXCOR1$(DES%)
  HASTA = FETEXCOR1$(HAS%)
  '
  Call ABREORFAB
  '
  SQLX$ = " SELECT COD_INTE, IDSUBOR, CanApro  FROM ORDEFABR WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE FechGen >= '" & DESDE & "' "
  SQLX$ = SQLX$ + " AND FechGen <= '" & HASTA & "' "
  SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC"
  '
'  Dim ORDEFABTMP As ADODB.Recordset
'  Set ORDEFABTMP = New ADODB.Recordset
'  ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  Set ORDEFABTMP = READSET(SQLX$)
  With ORDEFABTMP
    Do While Not .EOF
      CI% = XVALO(!COD_INTE)
      If ESTRAZABLE%(CI%) = 2 Then
         IDORFAB$ = SAFETEXT$(!IDSUBOR)
         KANAPRO& = XVALO(!CanApro)
         CONDICION$ = "IdSubOr = '" & IDORFAB$ & "' AND DepoNumse > 0 "
         KANTIREG& = XVALO(CantRegistros&("TANUMSE", CONDICION$))
         DIFERENCIA& = KANAPRO& - KANTIREG&
         If DIFERENCIA& <> 0 Then
           Call COMUNI("Orden de Fabricacion: " & IDORFAB$ & "    CODIGO: " & CODEXT$(CI%) & " \Cantidad O.F.:  " & KANAPRO& & "\Cantidad N.Serie :  " & KANTIREG&)
           '
           ' esta parte suprimida es para intentar la recuperacion.
           ' Realizado el 30/04 y dejado pendiente por EPB
'          NUSEMIN = VALOBADA("TANUMSE", "MIN(NUMEROSERIE)", "IDSUBOR = '" & IDORFAB$ & "'")
'47           If CantRegistros("TANUMSE", "NUMEROSERIE = " & (NUSEMIN - 1)) > 0 Then
'               NUSEMIN = NUSEMIN - 1
'               GoTo 47
'             End If
'          NUSEMAX = VALOBADA("TANUMSE", "MAX(NUMEROSERIE)", "IDSUBOR = '" & IDORFAB$ & "'")
'57           If CantRegistros("TANUMSE", "NUMEROSERIE = " & (NUSEMAX + 1)) > 0 Then
'               NUSEMAX = NUSEMAX - 1
'               GoTo 57
'             End If
'             '
         End If
      End If
    .MoveNext
    Loop
  End With
  Call COMUNI("Proceso Terminado")
  Command53.Enabled = True
End Sub

Private Sub Command54_Click()
    Command54.Enabled = False
    If DERACCE%("RECUFIFO", "Recuperación de Lotes") < 2 Then GoTo 99

    INTERRUMPE% = 0
    Label12.Visible = True
    Screen.MousePointer = 11
    Call ABRECONEXION
    SQLX$ = "SELECT COUNT(*) AS CANREGIS FROM MOVISTO "
    SQLX$ = SQLX$ + " INNER JOIN MASTER "
    SQLX$ = SQLX$ + " ON MOVISTO.Cod_Inte = MASTER.Codint "
    SQLX$ = SQLX$ + " WHERE MASTER.Trazable = '1' "
    SQLX$ = SQLX$ + " AND ( MOVISTO.CantIngre > 0 OR MOVISTO.CantSalid > 0 ) "
    SQLX$ = SQLX$ + " AND MOVISTO.CodiMovi = 'CF' "
    SQLX$ = SQLX$ + " AND ( MOVISTO.IdenLote = '' "
    SQLX$ = SQLX$ + " OR MOVISTO.IdenLote IS NULL )"
    Set CANRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
    FIN& = CANRE!CANREGIS
    CANRE.Close
    Set CANRE = Nothing
    '
    MsgBox "Registros Pendientes: " & FIN&
    FEMOVI = FECHANUM("01/01/09")
    '
6   FEMOVX$ = FETEXCOR1$(Int(FEMOVI))
    Label12 = FECHATEX$(Int(FEMOVI))
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " INNER JOIN MASTER "
    SQLX$ = SQLX$ + " ON MOVISTO.Cod_Inte = MASTER.Codint "
    SQLX$ = SQLX$ + " WHERE MASTER.Trazable = '1' "
    SQLX$ = SQLX$ + " AND ( MOVISTO.CantIngre > 0 OR MOVISTO.CantSalid > 0 ) "
    SQLX$ = SQLX$ + " AND MOVISTO.CodiMovi = 'CF' "
    SQLX$ = SQLX$ + " AND FECHMOV = '" & FEMOVX$ & "' "
    SQLX$ = SQLX$ + " AND ( MOVISTO.IdenLote = '' "
    SQLX$ = SQLX$ + " OR MOVISTO.IdenLote IS NULL )"
    '
    Dim MOVISELE As ADODB.Recordset
    Set MOVISELE = New ADODB.Recordset
    MOVISELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    Screen.MousePointer = 11
    With MOVISELE
       JJ& = 0: NUORAN% = 0
       Do While .EOF = False
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         ' Label12 = JJ&
         If INTERRUMPE% = 1 Then GoTo 39
         CIXI% = !COD_INTE
         FMOV = CVINT(!FECHMOV)
         If ESTRAZABLE%(CIXI%) <> 1 Then GoTo 29
         '
         On Error GoTo 15
         NOFA$ = TRIM$(!DoPriCor)
         CANCON = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
         '
         ' busca si hay uno para la misma O/F
         SQLY$ = "SELECT * FROM MOVISTO "
         SQLY$ = SQLY$ + " WHERE CodiMovi = 'CF' "
         SQLY$ = SQLY$ + " AND COD_INTE = " & CIXI%
         SQLY$ = SQLY$ + " AND DoPriCor = '" & NOFA$ & "' "
         SQLY$ = SQLY$ + " AND ( CantIngre > 0  or CantSalid > 0 ) "
         SQLY$ = SQLY$ + " AND IdenLote <> '' "
         SQLY$ = SQLY$ + " AND IdenLote IS NOT NULL "
         '
         Dim MOVILOT As ADODB.Recordset
         Set MOVILOT = New ADODB.Recordset
         Set MOVILOT = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         If Not (MOVILOT.BOF And MOVILOT.EOF) Then
            Do While Not MOVILOT.EOF
              IDLO$ = TRIM$(MOVILOT!IDENLOTE)
              If IDLO$ <> "" Then
                SLXX = SaldoLote(CIXI%, IDLO$)
                If SLXX >= CANCON Then
                  FALTA = CVINT(VALOBADA("LOTINGRE", "FechAlta", "IdenLote = '" & IDLO$ & "' AND Cod_Inte = " & CIXI%))
                  If FALTA <= FMOV Then
                     !IDENLOTE = Left$(IDLO$, 16)
                     .Update
                     Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
                     GoTo 29
                  End If
                End If
              End If
            MOVILOT.MoveNext
            Loop
         End If
         MOVILOT.Close
         Set MOVILOT = Nothing
         '
         ' BUSCA POR LISTA DE LOTES DISPONIBLES
         Call ARMALILOT(CIXI%)
         FINU& = ULTREG&("!LILOTSE") ' RECORRE EL LILOTSE
         '
         For U& = 1 To FINU&
            XX$ = REGLEIDO$("!LILOTSE", U&) ' REGISTRO LEIDO
            IDLOT$ = TRIM$(Mid$(XX$, 1, 16))
            If TRIM$(IDLOT$) <> "" Then
              SLXX = SaldoLote(CIXI%, IDLOT$)
              FALTA = CVINT(VALOBADA("LOTINGRE", "FechAlta", "IdenLote = '" & IDLOT$ & "' AND Cod_Inte = " & CIXI%))
              If FALTA <= FMOV Then
                If SLXX >= CANCON Then
                  !IDENLOTE = Left$(IDLOT$, 16)
' MsgBox CODEXT$(CIXI%) & "   " & IDLOT$
                  .Update
                  Call ACSALOTE(CIXI%, IDLOT$, "NOMESS")
                  GoTo 29
                End If
              End If
            End If
         Next U&
         '
         GoTo 29
         '
15       Resume 29
         '
29       On Error GoTo 0
       .MoveNext
       Loop
    End With
    '
39  MOVISELE.Close
    Set MOVISELE = Nothing
    FEMOVI = DIPOST(FEMOVI)
    If INTERRUMPE% < 1 Then
      If FEMOVI < HOY(HO$) Then
        GoTo 6
      End If
    End If
    '
    Screen.MousePointer = 1
99  Command54.Enabled = True
    '
End Sub

Private Sub Command55_Click()
'    OPX% = COMALTER%("Opciones de Trabajo con O/F's:\\ABM\ESTRUCTURA")
'    If OPX% = 1 Then
'       ABMCERTI09.Show 1
'    ElseIf OPX% = 2 Then
'      CERTIFICADOS.Show 1
'   End If
'   Exit Sub
   If CLACONTRA% = 1 Then
     SQLX$ = " select NUMEBUSQUE, NUMECOMPLE, REFERENCIA from pdoc_tbl "
     SQLX$ = SQLX$ + " where referencia like 'Orden de Fabricaci%' and tipodocu <>  'Etiqueta Orden Fabricacion' "
     SQLX$ = SQLX$ + " AND  tipodocu <>  'Cierre de O-Fabricacion'"
     SQLX$ = SQLX$ + " AND  tipodocu <>  'Remito de Traslado'"
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
25   On Error Resume Next
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0
      '
      JJ& = 0
      With RS
      Do While Not .EOF

       NUMEOR = XVALO(Mid$(SAFETEXT$(!REFERENCIA), 25, 6))
       !NUMEBUSQUE = NUMEOR
       NORFA$ = Mid$(SAFETEXT$(!REFERENCIA), 22)
       !NUMECOMPLE = NORFA$
       .Update
       .MoveNext
      Loop
     End With
End If
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
         j& = 0
         Screen.MousePointer = 11
         For U& = 1 To ULTREG&("ORDEFABR")
           x$ = REGLEIDO$("ORDEFABR", U&)
           STAT% = Asc(Mid$(x$, 108, 1))
           If STAT% > 0 Then
             If STAT% < 3 Then
               CIXI% = CVI(Mid$(x$, 83, 2))
               Call REPLA(x$, MKI$(CIXI%), 1, 2)
               Call REPLA(x$, MKI$(0), 3, 2)
               Call REPLA(x$, Space$(64), 5, 64)
               CATOT# = CVD(Mid$(x$, 89, 8))
               CAPRA# = CVD(Mid$(x$, 119, 8))
               CAPEN# = CATOT# - CAPRA#
               If CAPEN# < 0 Then CAPEN# = 0
               Call REPLA(x$, MKD$(CAPEN#), 159, 8)
               Call REPLA(x$, MKD$(0), 167, 8)
               j& = j& + 1
               Call GRAREG("ORFAMAT", x$, j&)
             End If
           End If
         
         Next U&
         Call SETULTREG("ORFAMAT", j&)
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
    OPROYECTA07.Show 1
'    Call ACTECOR(1)
'    Call GENEOPERFAP
'    Call FINAL("*OPERFAP")
'    Screen.MousePointer = 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
    '
50  TTOP$ = ""
    OPMAX% = 0
    OPTRAZA% = 0
    OPREPA% = 0
    OPSIMU% = 0
    GENNUM% = 0
    ENSAEQ% = 0
    If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
       TTOP$ = "Registro de Trazabilidad de O/F's"
       OPTRAZA% = 1
       OPMAX% = 1
    End If
    If EXISTE%("REPARA07.EXE") > 0 Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPREPA% = OPMAX%
       TTOP$ = TTOP$ + "Sub-Sistema de Reparaciones"
    End If
    ' '\\\OT-07-0386-FL-17/10/07///
    If EXISTE%("SIMULA07.EXE") > o Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPSIMU% = OPMAX%
       TTOP$ = TTOP$ + "Simulación de Restricciones de Stock"
    End If
    ' '\\\OT-07-0386-FIN///
    If EXISTE%("ENSABOMB.EXE") > o Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPENSABO% = OPMAX%
       TTOP$ = TTOP$ + "Ensayo de Bombas"
    End If
    '
    If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPTEROP% = OPMAX%
       TTOP$ = TTOP$ + "Servicios de Terceros"
    End If
    '
    If EXISTE%("TRANSE07.EXE") > 0 Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       ENSAEQ% = OPMAX%
       TTOP$ = TTOP$ + "Ensayo de Equipos - Control Final"
       '
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       GENNUM% = OPMAX%
       TTOP$ = TTOP$ + "Generador de Números de Serie"
    End If
    '
    If EXISTE%("RECHLI07.EXE") > 0 Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPRELI% = OPMAX%
       TTOP$ = TTOP$ + "Rechazo de Linea"
    End If
    '
    OPMAX% = OPMAX% + 1
    If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
    OPAJULOOF% = OPMAX%
    TTOP$ = TTOP$ + "Ajuste de Lote Con Asignación a O-F"
    '
    If InStr(UCase$(EMPREAC$), "MEDITEA") > 0 Then
        OPMAX% = OPMAX% + 1
        If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
        REEMPLAZO% = OPMAX%
        TTOP$ = TTOP$ + "Reemplazo de Componentes"
        '
        OPMAX% = OPMAX% + 1
        If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
        CANJECOMP% = OPMAX%
        TTOP$ = TTOP$ + "Canje de Componentes"
        '
        OPMAX% = OPMAX% + 1
        If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
        CANJEXTIPO% = OPMAX%
        TTOP$ = TTOP$ + "Canje de Componentes Por Tipo"
        
    End If
    '
    If InStr(UCase$(EMPREAC$), "ARCON") > 0 Then
        OPMAX% = OPMAX% + 1
        If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
        ABMCERTI% = OPMAX%
        TTOP$ = TTOP$ + "ABM de Certificados"
        '
        OPMAX% = OPMAX% + 1
        If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
        DATABOOK% = OPMAX%
        TTOP$ = TTOP$ + "Databook"
        '
    End If

    If TTOP$ <> "" Then
       TTOP$ = TTOP$ + "\Otros Accesos Directos"
    End If
    '
    If OPMAX% < 1 Then GoTo 20
    '
    OPX% = ALTERNA%(TTOP$)
    If OPX% > OPMAX% Then
         GoTo 20
       ElseIf OPX% < 1 Then
         GoTo 99
    End If
    '
    If OPX% = OPTRAZA% Then
        If DERACCE%("TRAZORF", "Registro de Trazabilidad de O/F's") = 2 Then
          ACONEC$ = "TRAZAB07"
          If EXISTE%("TRAZAB07.EXE") < 1 Then ACONEC$ = "TRAZAB08"
          Call CONECRUN(ACONEC$ + "/OFABRIC", "Registro de Trazabilidad de O/F's")
        End If
        GoTo 99
      ElseIf OPX% = OPREPA% Then
        If DERACCE%("REPARA", "Sub-Sistema de Reparaciones") = 2 Then
            ACONEC$ = "REPARA07"
            If EXISTE%("REPARA09.EXE") > 0 Then ACONEC$ = "REPARA09"
            Call CONECRUN(ACONEC$, "Sub-Sistema de Reparaciones")
        End If
        GoTo 99
      ElseIf OPX% = OPSIMU% Then
        If DERACCE%("SIMULA07", "Simulación de Restricciones de Stock") = 2 Then
          Call CONECRUN("SIMULA07", "Simulación de Restricciones de Stock")
        End If
        GoTo 99
      ElseIf OPX% = OPENSABO% Then
        Call CONECRUN("ENSABOMB", "Ensayos de Bombas")
        GoTo 99
      ElseIf OPX% = OPRELI% Then
        Call CONECRUN("RECHLI07", "Rechazo en Linea")
        GoTo 99
      ElseIf OPX% = OPTEROP% Then
        OPTEROP07.Show 1
        GoTo 99
      ElseIf OPX% = ENSAEQ% Then
        Call CONECRUN("TRANSE07/ENSAEQ", "Ensayos y Numeros de Serie")
        GoTo 99
      ElseIf OPX% = GENNUM% Then
        Call CONECRUN("TRANSE07", "Ensayos y Numeros de Serie")
        GoTo 99 'AGREGADO POR QUE ESTABA FALTNDO EN ESTA OPCION
      ElseIf OPX% = OPAJULOOF% Then
        Call AJUSTELOTEOF
        GoTo 99
      ElseIf OPX% = REEMPLAZO% Then  'SOLO MEDITEA
        Call REEMPLAZO_COMPONENTES
        GoTo 99
      ElseIf OPX% = CANJECOMP% Then  'SOLO MEDITEA
        Call CANJE_COMPONENTES
         GoTo 99
      ElseIf OPX% = CANJEXTIPO% Then  'SOLO MEDITEA
        Call CANJE_COMPONENTES_TIPO
      ElseIf OPX% = ABMCERTI% Then
         ABMCERTI09.Show 1
         GoTo 50
      ElseIf OPX% = DATABOOK% Then
         CERTIFICADOS.Show 1
         GoTo 50
    End If
    
    '
20  Call ACCDIR("ORDEFABR")
    '
99  Command8.Enabled = True
End Sub

Sub AJUSTELOTEOF()
    '
    If DERACCE%("AJUSLOOF", "Ajuste de Lote Con Asignación a O.Fabricación") < 2 Then Exit Sub
    '
    'MUESTRO OBJETO
    
10  OBX$ = OBJPLA$("SELDEPYOF", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    NORFA$ = TRIM$(Mid$(OBX$, 1, 12))
    DEPSEL% = Asc(Mid$(OBX$, 13, 1))
    '
    'VALIDACIONES
    If DEPSEL% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPSEL%))) = "" Then
      Call COMUNI("Depósito No Válido")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
      Call COMUNI("Orden de Fabricación No Válida\  \Recuerde que Solo se Pueden Ajustar\Ordenes de Fabricación en Curso Actual")
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    'MUESTRO FORMULARIO DE SELECCION DE LOTES
20  SELCODLOT07.Show 1
    '
    If TRIM$(SELCODLOT07.Text1) = "" And TRIM$(SELCODLOT07.Label28) = "" Then
       VDEF$ = OBX$
       GoTo 10
    End If
    
    'VALIDACIONES
    CIIXI% = CODINT(SELCODLOT07.Text1)
    If CIIXI% < 1 Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 20
    End If
    CONDI$ = "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CIIXI%
    COMPENOF% = CantRegistros&("RESERVA", CONDI$, "NOMESS")
    If COMPENOF% < 1 Then
       Call COMUNI("Código No Reservado Para : " & NORFA$)
       GoTo 20
    End If
    
    'VALIDO EL LOTE
    LOTSELEC1$ = TRIM$(SELCODLOT07.Label28)
    If LOTSELEC1$ <> "" Then
        'VERIFICO QUE EL ARTICULO ESTE RESERVADO PARA LA OF
        If ESTRAZABLE%(CIIXI%) = 1 Then
           CONDI1$ = "IdSubOr = '" & NORFA$ & "' AND IDENLOTE = '" & LOTSELEC1$ & "' AND Cod_Inte = " & CIIXI%
           LOTEASIG% = CantRegistros&("RESERVA", CONDI1$, "NOMESS")
           If LOTEASIG% < 1 Then
             Call COMUNI("Lote No Asignado Para : " & NORFA$)
             GoTo 20
           End If
        End If
    End If
    '
    'OBTENGO EL STOCKS PARA EL ARTICULO DEPOSITO Y OF
    CONDI2$ = "DoSecCor = '" & NORFA$ & "' AND Cod_Inte = " & CIIXI% & " AND DEPOMOVI = " & DEPSEL% & ""
    If ESTRAZABLE%(CIIXI%) = 1 Then ' SI ES TRAZABLE AG
       CONDI2$ = CONDI2$ + " AND IDENLOTE = '" & LOTSELEC1$ & "'"
    End If
    CANINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI2$, "NOMESS"))
    CANSALI = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI2$, "NOMESS"))
    STOCKX = CANINGRE - CANSALI
    '
    'PRESENTAR INPUTBOX CON LA CANTIDAD DE STOCKS
30  CANTSTOCK$ = TRIM$(InputBox$("Stocks", "Ingrese Cantidad Ajustada", Str(STOCKX)))
    If CANTSTOCK$ = "" Then GoTo 20
    '
    'VALIDAR QUE NO SEA MENOR LA CANTIDAD A AJUSTAR
    CANAJUS = XVALO(CANTSTOCK$)
    If CANAJUS <= STOCKX Then
      Call COMUNI("La Cantidad Ajustada Debe Ser Mayor a la Cantidad de Stocks")
      GoTo 30
    End If
    '
    'REALIZAR AJUSTE CON MOVISTOK
    CANTI = CANAJUS - STOCKX
    If CANTI > 0 Then
        DOCUNUM& = ProNroComprobante("AJ")
        DOPRILA$ = "AJ." & Str$(DOCUNUM&)
        DOPRICO$ = DOPRILA$
        DOSELA$ = NORFA$
        DOSECO$ = DOSELA$
        REFE$ = "Aj.Aut.Comp. " & NORFA$
        FF% = HOY(HO$)
        DEP% = DEPSEL%
        '
        Call MOVISTOK(FF%, CIIXI%, LOTSELEC1$, "AJ", DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, 1, 0, 1, 0, DEP%, CANTI, 0, 0, "", "", "", 0)
        Call COMUNI("Ajuste Realizado")
    End If
    '
    Screen.MousePointer = 1
    
End Sub
Private Sub Command9_Click()
    Command9.Enabled = False
    RESERV07.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Call ORFANU
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    OPX% = COMALTER%("Opciones de Trabajo con O/F's:\\Cierre de O/F's\Re-Apertura\Anular Cierre de O/F\Aprobacion de O-F")
    If OPX% = 1 Then
       If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
          If DERACCE%("CIERROFT", "Cierre de O/F's con Trazabilidad") = 2 Then
             ACONEC$ = "TRAZAB07"
             If EXISTE%("TRAZAB07.EXE") < 1 Then ACONEC$ = "TRAZAB08"
             Call CONECRUN(ACONEC$ + "/ORFCIE", "Cierre de O/F's con Trazabilidad")
          End If
          GoTo 99
        End If
        Call ORFCIE
      ElseIf OPX% = 2 Then
        If DERACCE%("REAPEROF", "Re-Apertura O/F's con Saldo") >= 2 Then
          Call REAPEROF
        End If
      ElseIf OPX% = 3 Then
        If DERACCE%("ANUCIEOF", "Anulación Cierre Parcial de O/F") >= 2 Then
          Call ANUCIEOF
        End If
      ElseIf OPX% = 4 Then
        If InStr(1, EMPREAC$, "ARCON") > 0 Then
            If DERACCE%("APCIEOF", "Aprobación Cierre Parcial de O/F") >= 2 Then
                Call APCIEOF
            End If
        Else
            Call COMUNI("Opcion No Disponible en la Presente Versión")
        End If
    End If
99  Command46.Enabled = True
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    Call REVIORFLAN
    If Option1.Value = True Then
       Screen.MousePointer = 11
10     Call GENORDEFDAT(0)
       Screen.MousePointer = 11
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
       '\\\OT-06-0045-WAR-09/02/06///
'           SQLX$ = SQLX$ + "where CanApro < 0.05 "
'           SQLX$ = SQLX$ + "AND CanRech < 0.05 "
       SQLX$ = SQLX$ + "WHERE StatuOrf < 3 " ' PARA QUE APAREZCAN LAS O/F CON CIERRE PARCIAL
       '\\\OT-06-0045-WAR-FIN///
       SQLX$ = SQLX$ + "order by IdSubOr ASC"
       'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
        Dim ORFATEMP As ADODB.Recordset
        Set ORFATEMP = New ADODB.Recordset
        ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

       '
       On Error Resume Next
       With ORFATEMP
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No Hay O/F's Pendientes\que Puedan Refrescarse.")
           Screen.MousePointer = 1
           On Error GoTo 0
           GoTo 99
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While Not .EOF
           If Val(Mid$(!IDSUBOR, 11)) = 1 Then
             ZZ$ = REGBLAN$("ORFANUE")
             REFXX$ = Left$(TRIM$(Left$(CODEXT$(!COD_INTE), 15)) + " - " + TRIM$(DESCRIT0$(!COD_INTE)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
             Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
             Call REPLA(ZZ$, REFXX$, 5, 64)
             JJ& = JJ& + 1
             Call GRAREG("!ORFANUE", ZZ$, JJ&)
           End If
           .MoveNext
         Loop
       End With
       Set ORFATEMP = Nothing
       Call SETULTREG("!ORFANUE", JJ&)
       Call CIERRARCH("!ORFANUE")
       Call FRESHECHO("!ORFANUE")
       Call TRACENTRAL("ORFANUE", "")
       Call FRESHECHO("ORFANUE")
       Screen.MousePointer = 1
       '
       Call CONECRUN("*ORF-C01", "Consulta por Número de O.F.")
       'Call GENORDEFDAT(0)
       'Call CONECRUN("*CONORFA", "Consulta Base de datos de O.F's")
    Else
       Call GENORDEFDAT(0)
       Call FINAL("*LISORFA")
    End If
99  Command61.Enabled = True
    Screen.MousePointer = 1
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
        '\\\OT-05-0671-WAR-12/10/05///
        VDEF$ = Space(5)
        Call REPLA(VDEF$, MKI$(0), 1, 1)
        Call REPLA(VDEF$, MKI$(COMESACT%), 2, 2)
          HOII% = HOY(HOS$)
        Call REPLA(VDEF$, MKI$(HOII%), 4, 2)
        '\\\OT-05-0671-WAR-FIN///
5       OPX$ = OBJPLA$("SELORFSTA", VDEF$)
        ''\\\OT-05-0671-WAR-12/10/05///
        DEIN% = CVI(Mid$(OPX$, 2, 2))
        HAIN% = CVI(Mid$(OPX$, 4, 2))
        FEX = DEIN%: DESX$ = FECHATEX$(FEX)
        FEX = HAIN%: HASX$ = FECHATEX$(FEX)
        PERIO1$ = "Desde: " + DESX$ + " Hasta: " + HASX$
        ''\\\OT-05-0671-WAR-FIN///
        If OPX$ <> "" Then
         TIPMAT% = Asc(Mid$(OPX$, 6, 1))
          If TIPMAT% > 0 Then
           If ECOARCH$("TATIMAT", Chr$(TIPMAT%)) = "" Then
            Call MENSERR(24, "Tipo de Material no Válido")
            GoTo 5
           End If
          End If
          '
          ESTADO$ = Mid$(OPX$, 1, 1)
          'If ECOARCH$("ORFASTA", OPX$) <> "" Then
          If ECOARCH$("ORFASTA", ESTADO$) <> "" Then
            DESI% = (-32767): HASI% = 32767: PERIO$ = ""
            If Asc(ESTADO$) = 3 Then
              HOII% = HOY(HOS$)
10            VDEF$ = MKI$(COMESACT%) + MKI$(HOII%)
              OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
              If OBX$ = "" Then GoTo 5
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
            j& = 0
            Screen.MousePointer = 11
            Call HEADERS("ORFASEL", "")
           ' Call HEADERS("ORFASEL", "Estado: " + ECOARCH$("ORFASTA", OPX$) + PERIO1$)
            Call HEADERS("ORFASEL", "Estado: " + ECOARCH$("ORFASTA", ESTADO$) + PERIO1$)
            '\\\OT-05-0671-WAR-12/10/05///
            If TRIM(PERIO$) <> "" Then
              Call HEADERS("ORFASEL", PERIO$)
            End If
            '\\\OT-05-0671-WAR-FIN///
            TIMAT$ = "TODOS"
            If TIPMAT% > 0 Then TIMAT$ = ECOARCH$("TATIMAT", Chr$(TIPMAT%))
            Call HEADERS("ORFASEL", "Tipo de Material: " + TIMAT$)
            '
            For U& = 1 To ULTREG&("ORDEFABR")
              x$ = REGLEIDO$("ORDEFABR", U&)
              'If Mid$(X$, 108, 1) = OPX$ Then 'ESTA LINEA ESTABA ASI
              TIPCOD% = TIMATE%(CVI(Mid$(x$, 83, 2)))
              If TIPMAT% = 0 Or TIPCOD% = TIPMAT% Then
               If Asc(Mid$(x$, 108, 1)) = Asc(Mid$(OPX$, 1, 1)) Then
                FECIX% = CVI(Mid$(x$, 117, 2))
                '\\\OT-05-0671-WAR-12/10/05///
                FEINX% = CVI(Mid$(x$, 81, 2))
                If FEINX% >= DEIN% Then
                 If FEINX% <= HAIN% Then
                '\\\OT-05-0671-WAR-FIN///
                  If FECIX% >= DESI% Then
                    If FECIX% <= HASI% Then
                      j& = j& + 1
                      Call GRAREG("ORFASEL", x$, j&)
                    End If
                  End If
                 End If
                End If
               End If
              End If
50          Next U&
            Call SETULTREG("ORFASEL", j&)
            Call CIERRARCH("*.*")
            Screen.MousePointer = 1
            '
            If j& < 1 Then
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
    If Option2.Value = True Then
10     Call CONCUMORFAB
    Else
       Call COMUNI("Solo Disponible Para Impresión")
       GoTo 10
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
Call RESERV07.Command7_Click
End Sub

Private Sub DepuBdD_Click()
Call RESERV07.Command5_Click
End Sub

Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    'Show
    'APLIXYZ$ = TRIM$(Command$)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture14)
    
    If FECHATEX$(HOY(HO$)) = "20/08/10" Then
       Command54.Visible = True
    End If
    '
    VERANTER$ = "ORDEFABR"

    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call ACTUENCASOL ' LE GRABA 0  A LOS VALORES NULL DE LA TABLA ENCASOLINT EN CAMPO RESERVA
    '
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    If APLINVO$ = "OPRESERV" Then
      RESERV07.Show 1
      Call FINAL("")
    End If
    '
    If APLINVO$ = "PRODUPRO" Then
      Option2.Value = True
      Call Command64_Click
    End If
    '
    If UCase$(Left$(APLINVO$, 6)) = "ORFNUE" Then
       APLIXYZ$ = APLINVO$
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
        COMMAND3.Enabled = False
        Label5.Enabled = False
      End If
    End If
    '
    MODOGABAL% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "GABAL") > 0 Then MODOGABAL% = 1
    If MODOGABAL% > 0 Then
        Call CREACAMPO("", "ORFAETI", "IDSUBOR", 10, 12, 1)
        Call CREACAMPO("", "ORFAETI", "NUORDEN", 3, 0)
        Call CREACAMPO("", "ORFAETI", "STATUS", 3, 0)
        Call CREACAMPO("", "ORFAETI", "USUARIO", 10, 32)
        Call CREACAMPO("", "ORFAETI", "FECHACIERRE", 8, 0)
        Call CREACAMPO("", "ORFAETI", "LOTERECE", 10, 12)
        Call CREACAMPO("", "ORFAETI", "MARBORRA", 3, 0)
    End If
    '
    
    Openforms = DoEvents
    Show
    '
    ' ESTO ES PARA QUE GENERE AUTOMÁTICAMENTE UNA O/F
    If APLIXYZ$ <> "" Then
       If UCase$(Left$(APLIXYZ$, 6)) = "ORFNUE" Then
          AppActivate Caption
          SendKeys Chr$(1)
          Call Command41_Click
          Exit Sub
       End If
    End If
    '
End Sub
'

Sub VALEMATER(MODOVA%)
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   If MODOVA% < 1 Or MODOVA% > 2 Then
      OPX% = ALTERNA%("Por Sector Productivo\Por O/Fabricacion\Devolución por O/F")
      If OPX% < 1 Or OPX% > 3 Then
          Exit Sub
        ElseIf OPX% = 2 Then
           Call VALEMATPOROF: Exit Sub
        ElseIf OPX% = 3 Then
           Call DEVOPOROF: Exit Sub
      End If
      If InStr(1, EMPREAC$, "GABAL") > 0 Then
           Call COMUNI("Opcion No Disponible en la Presente Version")
           Exit Sub
      End If
      MODOVA% = ALTERNA%("Vale de Preparacion\Vale de Transferencia")
      If MODOVA% < 1 Or MODOVA% > 2 Then Exit Sub
   End If
   '
   If MODOVA% = 2 Then
     If DERACCE%("VALETRAN", "Vale de Transferencia de Stocks") < 2 Then
       Exit Sub
     End If
   End If
   '
10 CAOSEL% = 0
   XX$ = REGSEL$("SECPROD/Sector Productivo Destino")
   If Len(XX$) <= 4 Then
      Exit Sub
   End If
   '
   SECCI = CVS(Left$(XX$, 4))
   SECPROD$ = Mid$(XX$, 5, 32)
   DEPODEST% = Asc(Mid$(XX$, 81, 1))
   If DEPODEST% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%))) = "" Then
     Call MENSERR(24, "Falta Definir Depósito Asociado\al Sector '" + TRIM$(SECPROD$) + "'.")
     GoTo 10
   End If
   '
   DepOrig% = Val(CONTROL$("", "DEPOCEN"))
   If DepOrig% < 1 Then
      'Call MENSERR(24, "Falta Definir Depósito Central")
      GoTo 13
   End If
   '
   EPAC$ = EMPREAC$
   If InStr(EPAC$, "DOSIVAC") > 0 Then
13    Call SELECHO("TADEPOSI/Depósito Origen de Materiales")
      DepOrig% = Val(VALACT1$("TADEPOSI"))
      If DepOrig% < 1 Then Exit Sub
   End If
   If DepOrig% = DEPODEST% Then
      Call MENSERR(24, "Operación Incorrecta.\Deposito Origen igual a Destino.")
      GoTo 10
   End If
   DEPOSI$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
   Call BLANARCH("!VALEMAT")
   '
   Screen.MousePointer = 11
   For U& = 1 To NUMAS%
      CANECE(U&) = 0
   Next U&
   '
   'SELECCIONAR ORDENES LANZADAS DE ESE SECTOR
   Call GENORDEFDAT(1, SECCI)
11 FIN& = ULTREG&("!ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         If CVS(Mid$(XX$, 113, 4)) = SECCI Then
            CAOSEL% = CAOSEL% + 1
            ORFASEL$(CAOSEL%) = Mid$(XX$, 69, 12)
         End If
      End If
   Next U&
   '
   'CARGAR MATERIALES RESERVADOS PARA O.F'S
   Call RESERV07.GENRESDAT
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
         If ORFA$ = ORFASEL$(KK%) Then
           GoTo 15
         End If
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
          STOCENT = STODEPOS(CI%, DepOrig%)
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
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + DEPOSI$ + " a " + SECPROD$
   VTB% = VENTABU%("VALEMAT" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CAITMO% = 0
   For YX& = 1 To ULTREG&("!VALEMAT")
      YY$ = REGLEIDO$("!VALEMAT", YX&)
      CI% = CVI(Mid$(YY$, 1, 2))
      CAN# = CVS(Mid$(YY$, 19, 4))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          If CAN# > 0.005 Then
            CAITMO% = CAITMO% + 1
            CIMOVIM%(CAITMO%) = CI%
            CAMOVIM(CAITMO) = (-1) * CAN#
          End If
        End If
      End If
   Next YX&
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
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
         REFE$ = "TR.a " + SECPROD$
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + DEPOSI$
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         XXX$ = REGBLAN$("TRANSFE")
         Call REPLA(XXX$, MKI$(FECHA%), 1, 2)
         Call REPLA(XXX$, Chr$(DepOrig%), 3, 1)
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
   CANCELPRINT% = 0
   IMPR$ = TRIM$(CONTROL$(IDENTER$, "SUIMPVAL"))
   If IMPR$ = "SI" Then CANCELPRINT% = 1
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
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + DEPOSI$
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

Sub DEVOPOROF()
   '
   If InStr(1, EMPREAC$, "GABAL") > 0 Then
        Call COMUNI("Opcion No Disponible en la Presente Version")
        Exit Sub
   End If
   
   Dim MATRANSFE#(32767)
   '
10 Call SELECHO("ECORDEP")
   '
   For t& = 1 To NUMAS%
      MATRANSFE#(t&) = 0
   Next t&
   '
   ARMAEXTER% = 0
   NUORFAB$ = VALACT1$("ECORDEP")
   If TRIM$(NUORFAB$) = "" Then Exit Sub
   '
   'Valida si es Armador Externo no permite devolucion
   If ValidaArmador%(NUORFAB$) > 0 Then
      Call COMUNI("Imposible Realizar Devolución !!!\Orden de Fabricación con Asignación de Armador  Externo")
      GoTo 10
   End If
   '
   VDEF$ = String$(64, 0)
   '
   CONDI$ = "IDSUBOR= '" & NUORFAB$ & "'"
   SPRD& = XVALO(VALOBADA("ORDEFABR", "SecciPro", CONDI$, "NOMESS"))
   For U& = 1 To ULTREG&("SECPROD")
     x$ = REGLEIDO$("SECPROD", U&)
     If CVS(Left(x$, 4)) = SPRD& Then
      DEST% = Asc(Mid$(x$, 81, 1))
      Exit For
     End If
   Next U&
   '
   If CONTROL$("", "DEPODEST") <> "" Then
     DEPORIGINAL% = CONTROL$("", "DEPODEST")
   End If
   '
   Call REPLA(VDEF$, NUORFAB$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPORIGINAL%), 13, 1) ' DEPOSITO TALLER
   '
12 OBX$ = OBJPLA$("SEORFADE", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   '
'   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DERACCE%("CAMBIDEP/NOMESS", "CAMBIO DE DEPOSITOS") < 2 Then
     If DEPORIGINAL% <> DEPODEST% Then
        Call MENSERR(24, "Usuario no Habilitado para Modificar Depósito\Módulo: CAMBIDEP")
        GoTo 12
     End If
   End If
   '
   'VALIDA DEPOSITOS
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then 'VALIDACION
     Call MENSERR(24, "Depósito Origen No Válido")
     VDEF$ = OBX$
     GoTo 12
   End If
   '
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then
      Call MENSERR(24, "Depósito Destino No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   Screen.MousePointer = 11
   '
   ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD INGRESADA POR ARTICULO SEGUN LA O.F.
   SQLX$ = "SELECT COD_INTE, CANTINGRE, CANTSALID FROM MOVISTO  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NUORFAB$ & "' "
   SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
   SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
'   Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set MoviTemp = READSET(SQLX$)
   With MoviTemp
     Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CAN = XVALO(!CANTINGRE) - XVALO(!CANTSALID)
       '
       MATRANSFE#(CI%) = MATRANSFE#(CI%) + CAN
       .MoveNext
     Loop
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD SALIDA POR ARTICULO SEGUN LA O.F.
   SQLX$ = "SELECT COD_INTE, CANTSALID, CANTINGRE  FROM MOVISTO  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF'"
   SQLX$ = SQLX$ + " AND DOPRICOR = '" & NUORFAB$ & "'"
   SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
'   Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set MoviTemp = READSET(SQLX$)
   With MoviTemp
     Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CAN = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
       MATRANSFE#(CI%) = MATRANSFE#(CI%) - CAN
      .MoveNext
     Loop
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
  '
   JJ& = 0
   Call SETULTREG("!MATEROF", 0)
   For I% = 1 To NUMAS%
     CI% = I%
     UM1$ = UNIMED$(CI%)
     CAN1 = MATRANSFE#(CI%)
     If MATRANSFE#(CI%) > 0.000005 Then
        JJ& = JJ& + 1
        y$ = REGBLAN$("!MATEROF")
        Call REPLA(y$, MKI$(CI%), 1, 2)
        Call REPLA(y$, UM1$, 3, 2)
        Call REPLA(y$, MKS$(CAN1), 5, 4)
        Call GRAREG("!MATEROF", y$, JJ&)
     End If
   Next I%
   If JJ& < 1 Then
      Call COMUNI("No se Encontró Material sin Consumir Para Esta Orden de Fabricación")
      Screen.MousePointer = 1
      GoTo 10
   End If
   
   Screen.MousePointer = 1
   VTB% = VENTABU%("MATEROF1")
   
   If VTB% < 0 Then GoTo 10
   
   OPPX% = COMALTER%("Confirmación de Devolución\\Realizar Devolución\Cancelar")
   If OPPX% < 1 Or OPPX% > 1 Then
       GoTo 10
   End If
   '
   'Devolucion de Materiales
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   HOII% = HOY(HOS$)
   FECHA% = HOII%
   LOTE$ = NUORFAB$
'   CMOV$ = "TR"
   DOSEC$ = "TR-" + TRIM$(Mid$(NUORFAB$, 4))
   NRO& = ProNroComprobante(CMOV$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   VUNI# = 0
   MONE$ = ""
   DOSEC$ = TRIM$(NUORFAB$) ' NUMERO DE ORDEN DE FABRICACION
   While Len(DOSEC$) < 10
      DOSEC$ = DOSEC$ + " "
   Wend

   'Recorro el Materof

   FIN& = ULTREG&("!MATEROF")
   For H& = 1 To FIN&
      YY$ = REGLEIDO$("!MATEROF", H&)
      CI% = CVI(Mid$(YY$, 1, 2))
      CAN = CVS(Mid$(YY$, 5, 4))
      If CAN > 0.05 Then
         LOTE$ = ""
         If ESTRAZABLE(CI%) = 1 Then
           CONDI$ = " IdSubOr = '" & NUORFAB$ & "' AND Cod_Inte = " & CI% & ""
           LOTE$ = VALOBADA("RESERVA", "IdenLote", CONDI$, "NOMESS")
         End If
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
          '
         Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CAN)
          '
         CAN = (-1) * CAN
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
         Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CAN)
         
      End If
   Next H&
   '
   Screen.MousePointer = 1
   Call COMUNI("Proceso Terminado")
   '
End Sub


Sub VERORFA()
    '
    Call ABREORFAB
    '
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        OPX% = COMALTER%("Opciones de Impresion\\O-Fabricacion\Etiqueta Existente\Nueva Etiqueta")
        If OPX% < 1 Or OPX% > 3 Then GoTo 99
        '
        If OPX% = 1 Then GoTo 3
        If OPX% = 2 Then GoTo 10
        If OPX% = 3 Then GoTo 20
    End If
    '
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    NPX& = XVALO(Mid$(NORFA$, 4, 6))
    TIDOCUM$ = "Orden de Fabricación"
    DOCUNU& = 0
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NORFA$)
    GoTo 3

10  If DERACCE%("REIETIOF", "Re-impresion Etiqueta de O-F") < 2 Then GoTo 99
11  Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    NUMEET$ = InputBox("Numero de Etiqueta", NORFA$, "")
    NN% = InStr(1, NUMEET$, "-")
    If NN% > 0 Then
        CO% = XVALO(Mid$(NUMEET$, 1, NN% - 1)):  FI% = XVALO(Mid$(NUMEET$, NN% + 1))
        If FI% < CO% Then
            Call MENSERR(24, "La Cantidad de Comienzo Debe Ser Mayor a la Final")
            GoTo 11
        End If
        For II% = CO% To FI%
            Call PRINTETIQUETASGABAL(NORFA$, II%)
        Next II%
    Else
        NUMEETI% = XVALO(NUMEET$): If NUMEETI% < 1 Then GoTo 11
        Call PRINTETIQUETASGABAL(NORFA$, NUMEETI%)
    End If
    GoTo 11

20  If DERACCE%("NUETIOF", "Genera Nueva Etiqueta de O-F") < 2 Then GoTo 99
21  Call SELECHO("ECORDEP/Indice General de Ordenes de Fabricación (Pendientes)")
    NORFA$ = TRIM$(VALACT1$("ECORDEP"))
    If NORFA$ = "" Then GoTo 99
    STAT% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IDSUBOR='" & NORFA$ & "'"))
    If STAT% > 1 Then
        Call COMUNI("Imposible Generar una Nueva Etiqueta.\O-F Cerrada o Anulada")
        GoTo 99
    End If
    Call GRABAETIQUETASGABAL(NORFA$, 2)
    NUMEETI% = XVALO(VALOBADA("ORFAETI", "MAX(NUORDEN)", "IDSUBOR='" & NORFA$ & "'"))
    Call PRINTETIQUETASGABAL(NORFA$, NUMEETI%)
    GoTo 21
    
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

Function ContRepetCompon%()
    '
    Repet% = 0 ' 0 no hay repeticion - 1 hay repeticion de componentes
    Dim YARESE%(32767) ' PONGO EN 0 TODOS LOS ELEMENTOS
    For I% = 1 To NUMAS%
       YARESE%(I%) = 0
    Next I%
    '
    FIN& = ULTREG&("!RESERVA")
    For H& = 1 To FIN&
       XX$ = REGLEIDO$("!RESERVA", H&)
       CI% = CVI(Mid$(XX$, 1, 2))
       If CI% > 0 Then
         If ESTRAZABLE%(CI%) < 1 Then   ' SOLO CONTROLA LOS NO TRAZABLES
           YARESE%(CI%) = YARESE%(CI%) + 1
           If YARESE%(CI%) > 1 Then
              Repet% = 1
              If COMALTER%("Código: '" & CODEXT$(CI%) & "' (no trazable) Repetido en Fórmula\\Cancelar Orden de Fabricación\Generar Igual") <> 2 Then
                Exit For
              End If
              Repet% = 0
           End If
         End If
       End If
    Next H&
    '
    ContRepetCompon% = Repet%
    '
End Function

Private Sub Frame9_DblClick()
  GENORF07.Show 1
End Sub

'
Private Sub GenReservOC_Click()
Call RESERV07.COMMAND3_Click
End Sub

Private Sub Image2_DBLClick()
  '\\\OT-05-0733-WAR-14/11/05///
  'NUEORFA04.Show 1
  '\\\OT-05-0733-WAR-FIN///
End Sub

Private Sub InfoPro_Click()
Call Command51_Click
End Sub

Private Sub Label12_DblClick()
    INTERRUMPE% = 1
End Sub

Private Sub LanOrd_Click()
Call Command42_Click
End Sub

Private Sub LiToImpre_Click()
'
RESERV07.Option6 = True
RESERV07.Option4 = True
Call RESERV07.Command2_Click
'

End Sub

Private Sub LiToPan_Click()
'
RESERV07.Option6 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
'
End Sub



Private Sub MaestArt_Click()
   '
       Call CONECRUN("ARCHIG07/AMASTO", "Padron Maestro de Articulos")
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
RESERV07.Option3 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
'

End Sub

Private Sub ReXNumOF_Click()
'
RESERV07.Option2 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
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

Sub ANUCIEOF()
   '
   MODOGABAL% = 0
   If InStr(1, EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
   If MODOGABAL% = 1 Then
      OPXX% = COMALTER%("Elija la Modalidad de Anulación de Cierre\\Código de Barras\Manual")
      If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
      If OPXX% = 2 Then GoTo 10
      '
      ANUORFA.Caption = "Anulación Cierre de O-Fabricación"
      ANUORFA.Label2 = "1"
      ANUORFA.Show 1
      Exit Sub
   End If
   '
   ' 1.- Selección del Producto
10 Call SELECHO("MASTER")
   If VALACT1$("MASTER") = "" Then Exit Sub
   CODD$ = VALACT1$("MASTER")
   CIXI% = CODINT%(CODD$)
   If CIXI% < 1 Then Exit Sub
   '
   If ESTRAZABLE%(CIXI%) = 2 Then ' SI ES TRAZABLE NO SE PUEDE ANULAR
      Call COMUNI("Codigo: " & CODEXT$(CIXI%) & " Trazable Nivel 2\Imposible Anular Cierre de O/F.")
      GoTo 10
   End If
   
   ' 2.- Selección de la O/F
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT DoPriCor, FECHMOV, CANTINGRE, FeReMovi  FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND CodiMovi = 'OF' "
   SQLX$ = SQLX$ + "AND CantIngre > 0 "
   SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC "
   'Set Movisel = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim MOVISEL As ADODB.Recordset
'   Set MOVISEL = New ADODB.Recordset
'   MOVISEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set MOVISEL = READSET(SQLX$)
   '
   With MOVISEL
      On Error Resume Next
      .MoveFirst '
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Cierres de O/F Reversibles\para el Código Elegido.")
         GoTo 10
      End If
      On Error GoTo 0
      '
      JJ& = 0
      Do While .EOF = False
         y$ = REGBLAN("ECORDEP")
         NOFA$ = !DoPriCor
         FERE$ = FECHATEX$(CVINT(!FECHMOV))
         CANX$ = FORMATNUM$(!CANTINGRE, "F10.1")
         If Not IsNull(!FeReMovi) Then DIVE$ = FORMATNUM$(CDbl(!FeReMovi) - 0.00001, "F18.6")
         TTXX$ = FERE$ + "    " + CANX$ + "  " + DIVE$
         Call REPLA(y$, NOFA$, 1, 12)
         Call REPLA(y$, TTXX$, 13, 52)
         '
         JJ& = JJ& + 1
         Call GRAREG("!ECORDEP", y$, JJ&)
         '
      .MoveNext
      Loop
   End With
   MOVISEL.Close
   Set MOVISEL = Nothing
   '
   Call SETULTREG("!ECORDEP", JJ&)
   Call CIERRARCH("!ECORDEP")
   Call FRESHECHO("!ECORDEP")
   Screen.MousePointer = 1
   '
   RSEL$ = REGSEL$("!ECORDEP")
   If Len(RSEL$) <= 4 Then GoTo 10
   '
   NOFASE$ = TRIM$(Left$(RSEL$, 12))
   CANUL = XVALO(Mid$(RSEL$, 25, 10))
   If COMALTER%("¿ Confirma la Anulación de Este Cierre ?\\No, Conservarlo\Si, Anularlo") <> 2 Then GoTo 10
   '
   Screen.MousePointer = 11
   FREMOX1 = XVALO(TRIM$(Mid$(RSEL$, 37, 18)))
   FREMOX21 = CDate(FREMOX1)
   FREMOX3% = CVINT(FREMOX21)
   FREMOX$ = FETEXCOR1$(FREMOX3%)
   FREMOX22 = CDate(FREMOX1 + 0.0002)   ' ESTO SON 17 SEGUNDOS
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE DoPriCor = '" & NOFASE$ & "' "
   SQLX$ = SQLX$ + "AND FeReMovi >= '" & FREMOX$ & "' "
   '
   Dim MOVISELE As ADODB.Recordset
   Set MOVISELE = New ADODB.Recordset
   MOVISELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   Screen.MousePointer = 11
   With MOVISELE
      On Error Resume Next
      .MoveFirst '
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Movimientos de Cierre Reversibles\para la O/F Elegida.")
         GoTo 10
      End If
      On Error GoTo 0
      '
      JJ& = 0: NUORAN% = 0
      Do While .EOF = False
         FEREMO = CDate(CDbl(!FeReMovi))
         If FEREMO < FREMOX21 Or FEREMO > FREMOX22 Then GoTo 29
         REFA$ = !ReferCor
         CIXI% = !COD_INTE
         NUORAN% = !NUORITEM
         !ReferCor = Left$("Anul." + REFA$, 64)
         !CANTINGRE = 0
         !CANTSALID = 0
         .Update
         If CIXI% > 0 Then
            CIMOVI%(CIXI%) = 1
            IDLOTX$ = TRIM$(!IDENLOTE)
            If IDLOTX$ <> "" Then
               Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS")
            End If
         End If
29    .MoveNext
      Loop
   End With
   '
39 Set MOVISELE = Nothing
   '
'3.- Actualiza saldo pendiente de la Orden de Fabricacion
   Call ABREORFAB
   SQLX$ = "SELECT * FROM OrdeFabr "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFASE$ & "'"
   Dim ORDEFFFAB As ADODB.Recordset
   Set ORDEFFFAB = New ADODB.Recordset
   ORDEFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With ORDEFFFAB
     On Error Resume Next
     If Not (.EOF And .BOF) Then
      On Error GoTo 0
      !CanApro = !CanApro - CANUL
      !StatuOrf = 1                ' MARCA COMO EN CURSO
      .Update
     End If
     On Error GoTo 0
   End With
   ORDEFFFAB.Close
   Set ORDEFFFAB = Nothing
   '
   Screen.MousePointer = 1
   '
   Call COMUNI("Movimiento Anulado")
   GoTo 10
   
End Sub

Sub REAPEROF()
    '
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        Call COMUNI("Opcion No Disponible en la Presente Version")
        Exit Sub
    End If
    
    REBLA$ = REGBLAN$("ECORDEP")
    '
    Call ABREORFAB
    SQLX$ = " SELECT DISTINCT IdSubOr, StatuOrf, CanApro, CanProy FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE StatuOrf = 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
        y$ = REBLA$
        Call REPLA(y$, !IDSUBOR, 1, 12)
        Call REPLA(y$, ECOARCH$("ECORDEF", !IDSUBOR), 13, 52)
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", y$, JJ&)
      .MoveNext
      Loop
    End With
    Set ORDEFABTMP = Nothing
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
        '.FindFirst "Idsubor = '" & NOFA & "' AND StatuOrf = 3 "
        SQLX$ = "SELECT Statuorf FROM OrdeFabr "
        SQLX$ = SQLX$ + " WHERE Idsubor = '" & NOFA & "' AND StatuOrf = 3"
        Dim ESTADORDEN As ADODB.Recordset
        Set ESTADORDEN = New ADODB.Recordset
        ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

        With ESTADORDEN
          On Error Resume Next
          'If .NoMatch = False Then
            '.Edit
            .MoveFirst
            If Err < 1 Then
              Do While Not .EOF
                !StatuOrf = 1
                .Update
                y$ = REBLA$
                Call REPLA(y$, NOFA$, 1, 12)
                Call REPLA(y$, ECOARCH$("ECORDEF", NOFA$), 13, 52)
                Call FILTERPUTRECORD("ECORDEP", 1, NOFA$, y$)
                PUDOABRIRSE% = 1
                '.FindNext "Idsubor = '" & NOFA$ & "' AND StatuOrf = 3 "
                .MoveNext
              Loop
            End If
            On Error GoTo 0
        End With
        Set ESTADORDEN = Nothing
        If PUDOABRIRSE% = 0 Then
            Call COMUNI("No Fue Posible Re-Abrir esta O/F")
          ElseIf PUDOABRIRSE% = 1 Then
            Call COMUNI(NOFA$ + " Re-Abierta.\No se Re-Generaron Reservas ni Operaciones")
        End If
      End If
    End If
    '
99  'OrdeFabr.Close
    '
End Sub
'
Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "GENORFAB"
           Call Command41_Click                 ' GENERA ORDEN DE FABRICACION
       Case "LANORFAB"
           Call Command42_Click                 ' LANZAMIENTO DE ORDEN DE FABRICACION
       Case "CIEORFAB"
           Call Command46_Click                 ' CIERRE DE ORDEN DE FABRICACION
       Case "SUSORFAB"
           Call Command44_Click                 ' SUSPENCION DE ORDEN DE FABRICACION
       Case "ANUORFAB"
           Call Command45_Click                 ' ANULACION DE OREDEN DE FABRICACION
       Case "MODPRISECPRO"
           Call Command43_Click                 ' MODIFICAR PRIORIDAD Y SECTOR PRODUCTIVO
       Case "REFORFAB"
           Call Command4_Click                  ' REFRESCAR ORDEN DE FABRICACION
       Case "ANOTAORFA"
           Call Command49_Click                 ' ANOTADOR DE ORDEN DE FABRICACION
       Case "CONGESRE"
           Call Command9_Click                  ' CONTROL DE GESTION DE RESERVAS
       Case "TRANSMAT"
           Call Command11_Click                 ' TRANSFERENCIA DE MATERIALES

       '
       'CONSULTAS Y LISTADOS
       '
       Case "CONORFAB"
           Option2.Value = True
           Call Command61_Click                 ' CONSULTA GENERAL POR ORDEN DE FABRICACCION
       Case "PENSECPRO"
           Option2.Value = True
           Call Command10_Click                 ' PENDIENTE POR SECTOR PRODUCTIVO
       Case "LISELESTA"
           Option2.Value = True
           Call Command62_Click                 ' LISTADO SELECTIVO POR ESTADO
       Case "PENCODPRO"
           Option2.Value = True
           Call Command64_Click                 ' PENDIENTE POR CODIGO DE PRODUCTO
       Case "CUMPORFAB"
           Option2.Value = True
           Call Command63_Click                 ' CONTROL DE CUMPLIMIENTO DE ORDEN DE FABRICACION
       Case "VERORFAB"
           Call Command18_Click                 ' VER Y REIMPRIMIR ORDEN DE FABRICACION
       Case "OPRORFA"
           Call Command7_Click                  ' OPERACIONES PROYECTADAS
       Case "CARTOTEQUI"
           Call Command6_Click                  ' CARGA TOTAL POR EQUIPO
       '
       'CONFIGURACION ORDEN DE FABRICACION
       '
       Case "CONGENORFAB"
           Call Command14_Click                  ' CONFIGURACION GENERAL DE ORDEN DE FABRICACION
      
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub ACTUENCASOL() ' ACTUALIZA EL CAMPO NUME_SOLINT PONIENDO EN 0 LOS CAMPOS NULL
    '
    Call ABRESTOCKS
    CONDI$ = " NUME_SOLINT IS NULL"
    '
    Call ACTUREGISTRO("RESERVA", "NUME_SOLINT", CONDI$, 0, REGAF&)
    '
End Sub

Sub VALEMATPOROF()
   '
   If InStr(1, EMPREAC$, "GABAL") > 0 Then
        Call TRANSFEGABAL
        Exit Sub
   End If
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   If RECLEN("VALEMAT") < 64 Then
      Call MENSERR(24, "Longitud de Registro de Tabla 'VALEMAT'\Debe ser 64 pos. Actual es de 32 pos.")
      Call FINAL("")
   End If
   '
 ARCHIVOX1$ = "VALEMAT"
' EPAC$ = TRIM$(UCase$(EMPREAC$))
' If InStr(EPAC$, "MEDITEA") > 0 Then ARCHIVOX1$ = "VALEMAT9"
   
   MODOGABAL% = 0
   If InStr(1, EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
   '
10 Call SELECHO("ECORDEP")
   NUORFAB$ = VALACT1$("ECORDEP")
   If TRIM$(NUORFAB$) = "" Then Exit Sub
   '
   ' SI NO TIENE ASIGNADO UN ARMADOR SIGUE CON PROCESO NORMAL
   If ValidaArmador%(NUORFAB$) > 0 Then
       CONDI$ = "IDSUBOR = '" & NUORFAB$ & "'"
       ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
       Call TRANSF_A_ARMADOR(NUORFAB$, ARMAD%)
       GoTo 10          ' PROXIMA ASIGNACION
   End If
   '
11 For U& = 0 To 32767:    CANECE(U&) = 0:   Next U&
   SPRD& = XVALO(VALOBADA("ORDEFABR", "SecciPro", "IDSUBOR= '" & NUORFAB$ & "'", "NOMESS"))
   For U& = 1 To ULTREG&("SECPROD")
      x$ = REGLEIDO$("SECPROD", U&)
      If CVS(Left(x$, 4)) = SPRD& Then
         DEST% = Asc(Mid$(x$, 81, 1))
         Exit For
      End If
   Next U&
   '
   If CONTROL$("", "DEPOCEN") <> "" Then DEPOCENTRAL% = XVALO(CONTROL$("", "DEPOCEN")) ' TRAE SEGUN CONFIGURACION
   '
   If CONTROL$("", "DEPODEST") <> "" Then DEPODESTINO% = XVALO(CONTROL$("", "DEPODEST"))
   If DEPODESTINO% < 1 Then DEPODESTINO% = DEST% 'SI NO TIENE CONFIGURADO
   '
   VDEF$ = String$(64, 0)
   Call REPLA(VDEF$, NUORFAB$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPODESTINO%), 14, 1)
   Call REPLA(VDEF$, Chr$(DEPOCENTRAL%), 13, 1)
   '
12 OBJETO$ = "SEORFADE"
   If MODOGABAL% = 1 Then OBJETO$ = "SEORFADG"
   OBX$ = OBJPLA$(OBJETO$, VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   '
   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DepOrig% <> DEPOCENTRAL% Or DEPODEST% <> DEPODESTINO% Then
      If DERACCE%("CAMBIDEP", "Cambiar Depositos de Transferencia") < 2 Then GoTo 12
   End If
   '
   'VALIDA DEPOSITOS
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then 'VALIDACION
      Call MENSERR(24, "Depósito Destino No Válido"): GoTo 12
   End If
   '
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then
      Call MENSERR(24, "Depósito Origen No Válido"): GoTo 12
   End If
   '
   Call ABRESTOCKS
   'RECORRE LA TABLA RESERVAS Y ARMA EL ARCHIVO RESERVAS
   CONDI$ = "CantRes - CantCons > 0 "
   CONDI$ = CONDI$ + " AND IDSUBOR = '" & NUORFAB$ & "'"
   '
   FIN& = CantRegistros("RESERVA", CONDI$)
   If FIN& < 1 Then
       Call MENSERR(24, " No Existen Reservas Efectuadas.\para esta Orden de Fabricacion.")
       GoTo 12
   End If
   
   SQLX$ = " SELECT IDSUBOR, COD_INTE, FechRes, CantRes, CANTCONS, Saldo_Entre, IdenLote, Deposito, UsoUnit  "
   SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) WHERE " + CONDI$
'   Dim RESERTMP As ADODB.Recordset
'   Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set RESERTMP = READSET(SQLX$)
   JJ& = 0: CONTA% = 0
   REBLA$ = REGBLAN("!RESERVA")
   Call SETULTREG("!RESERVA", 0)
   With RESERTMP
      Do While (.EOF = False)
          y$ = REBLA$
          CII% = XVALO(!COD_INTE)
          If ESTRAZABLE(CII%) = 2 Then
              CONTA% = CONTA% + 1 ' valida si hay codigos trazables
            Else
              Call REPLA(y$, MKI$(!COD_INTE), 1, 2)
              Call REPLA(y$, !IDSUBOR, 3, 12)
              Call REPLA(y$, MKI$(!COD_INTE), 15, 2)
              Call REPLA(y$, MKI$(CVINT(!FechRes)), 17, 2)
              Call REPLA(y$, MKD$(!CANTRES), 19, 8)
              Call REPLA(y$, MKD$(!CANTCONS), 27, 8)
              Call REPLA(y$, MKD$(!Saldo_Entre), 35, 8)
              Call REPLA(y$, SAFETEXT$(!IDENLOTE), 43, 16)
              Call REPLA(y$, Chr$(!Deposito), 60, 1)
              Call REPLA(y$, MKS$(!USOUNIT), 75, 4)
              '
              JJ& = JJ& + 1
              Call GRAREG("!RESERVA", y$, JJ&)
          End If
13    .MoveNext
      Loop
   End With
   RESERTMP.Close
   Set RESERTMP = Nothing
   Call SETULTREG("!RESERVA", JJ&)
   Call CIERRARCH("!RESERVA")
   '
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\los cuales Seran Excluidos de la Lista de Componentes")
   End If
    '
   FIN& = ULTREG&("!RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!RESERVA", U&)
      ORFA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) '- CVD(Mid$(XX$, 27, 8))
      '
      If CANTI < 0 Then CANTI = 0
      If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then CANTI = CANTI / 1000
      '
      IDLO$ = TRIM$(Mid$(XX$, 43, 16))
      CONDI$ = " DoSecLar = '" & TRIM$(ORFA$) & "'"
      CONDI$ = CONDI$ & " AND Cod_Inte = " & CI%
      CONDI$ = CONDI$ & " AND DepoMovi = " & DEPODEST%
      CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "'"
      CANTXX = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
      CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
      STODEST = CANTXX - CANTSAL 'STODEPOS(CI%, DEPODEST%)
      ' STOCENT = STODEPOS(CI%, DepOrig%)
      STOCENT = (SaldoLote(CI%, IDLO$, DepOrig%))
      If STOCENT < 0 Then STOCENT = 0
      NECNETA = CANTI - STODEST
      If NECNETA < 0 Then NECNETA = 0
      '
      TRANSFAHORA = 0
      If MODOGABAL% = 1 Then
        'CALCULO DEL SOBRANTE
        RESERVAS = CALCURESERVAINTE(NUORFAB$, CI%, DEPODEST%)
        STODEST = STODEPOS(CI%, DEPODEST%)
        NECNETA = RESERVAS
      End If
      '
      VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
      If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
         VALEMATX = STOCENT
      End If
      '
      If VALEMATX < 0.05 Then VALEMATX = 0
       '
      ZZ$ = REGBLAN$(ARCHIVOX1$)
      Call REPLA(ZZ$, MKI$(CI%), 1, 2)
      Call REPLA(ZZ$, MKS$(CANECE(CI%)), 3, 4) ' NECESIDAD BRUTA
        If MODOGABAL% = 1 Then Call REPLA(ZZ$, MKS$(CANTI), 3, 4)
      Call REPLA(ZZ$, MKS$(STODEST), 7, 4)     ' STOCK DEPOSITO DESTINO
      Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
      Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
      Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
      Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
      Call REPLA(ZZ$, "......", 27, 6)
      Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
      Call REPLA(ZZ$, IDLO$, 49, 16)
      Call GRAREG("!" & ARCHIVOX1$, ZZ$, U&)
   Next U&
   Call SETULTREG("!" & ARCHIVOX1$, FIN&)
   Call CIERRARCH("!" & ARCHIVOX1$)
   '
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", Chr$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
   VTB% = VENTABU%("VALEMAT9" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CONTA% = 0
   Archivx$ = "!" & ARCHIVOX1$
   For JJ1& = 1 To ULTREG&(Archivx$) ' valida si hay codigos trazables
      XX$ = REGLEIDO$(Archivx$, JJ1&)
      CI% = CVI(Mid$(YY$, 1, 2))
      If ESTRAZABLE(CI%) = 2 Then
        CONTA% = CONTA% + 1
      End If
   Next JJ1&
   '
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Deben ser Excluidos de la Lista de Componentes")
      GoTo 25
   End If
   '
   If CONTROL("VANEDECE", "DEPOCEN") = "SI" Then ' SI CONFIGURACION DE CONTROL DE NEGATIVOS
      For YX& = 1 To ULTREG&(Archivx$)
         YY$ = REGLEIDO$(Archivx$, YX&)
         CI% = CVI(Mid$(YY$, 1, 2))
         CAN# = CVS(Mid$(YY$, 19, 4))
             If CAN# > 0.005 Then
               STOCXCENT# = STODEPOS(CI%, DepOrig%) 'CVS(Mid$(YY$, 15, 4))
               VALMAT# = CVS(Mid$(YY$, 19, 4))
               If STOCXCENT# < VALMAT# Or STOCXCENT# < 0 Then
                 Call MENSERR(24, "Transferencia No Procesable\Código: " & CODEXT$(CI%) & "\Generará Un Stock Negativo en Depósito de Origen")
                 GoTo 25
               End If
             End If
      Next YX&
   End If
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
'   Call ASIGNALOTE_RESERVA_VALEMAPOROF(NUORFAB$, Archivx$, OK%)
'   If OK% < 0 Then GoTo 25
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then GoTo 10
   
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   NRO& = ProNroComprobante(CMOV$)
   '
   Call BLOQARCH("MOVISTO")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   'LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   'DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
   DOSEC$ = TRIM$(ORFA$) ' NUMERO DE ORDEN DE FABRICACION
    While Len(DOSEC$) < 10
      DOSEC$ = DOSEC$ + " "
    Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&(Archivx$)
      XX$ = REGLEIDO$(Archivx$, U&)
      CI1% = CVI(Mid$(XX$, 1, 2))
      If CVS(Mid$(XX$, 19, 4)) >= FRACMIN(CI%) Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         '
         LOTE$ = TRIM$(Mid$(XX$, 49, 16)) ' INICIALIZAMOS POR SI NO ENTRA NO QUEDE ANTERIOR
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
         '
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
      End If
   Next U&
   '
   Call LIBARCH("MOVISTO")
   
   Call BAJALDISCO
   '
   'IMPRIME EL VALE DE TRANSFERENCIA
   '
   CANCELPRINT% = 0
   IMPR$ = TRIM$(CONTROL$(IDENTER$, "SUIMPVAL"))
   If IMPR$ = "SI" Then CANCELPRINT% = 1
   '
60 FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(Str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
      CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUORFAB$)
      CAPARDOC% = 3
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CODTABU1$(6) = "REF-MAT1"
      CACOLTAB1% = 6
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For U& = 1 To ULTREG&(Archivx$)
         XX$ = REGLEIDO$("!" & ARCHIVOX1$, U&)
         CICRU% = CVI(Mid$(XX$, 1, 2))

         If CVS(Mid$(XX$, 19, 4)) >= FRACMIN(CICRU%) Then
            '
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            KICOL1$ = Str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = UNIMED$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
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
   Call COMUNI("Transferencia Realizada")
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub
Sub ASIGNALOTE_RESERVA_VALEMAPOROF(NORFA$, ARCHIVO1$, OK%)

      '
'SE CAMBIO LA CNTIDD DE DECIMLES QUE ERA FIJA DE 1  X  & CNTDC$(CI%)
'ESTO EN ADMIORF07 (MODUILO)
    OK% = 0
    TRAZCOMPO.List2.Clear
    For YX& = 1 To ULTREG&(ARCHIVO1$)
       YY$ = REGLEIDO$(ARCHIVO1$, YX&)
       CI% = CVI(Mid$(YY$, 1, 2))
       CANTI = CVS(Mid$(YY$, 19, 4))
       LOTE$ = Mid$(YY$, 49, 16)
       ESTRAZ% = ESTRAZABLE%(CI%)
       If ESTRAZABLE%(CI%) = 1 Then
         If CANTI > FRACMIN(CI%) Then
            TX$ = ""
            Call REPLA(TX$, CODEXT$(CI%), 1, 16)
            Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CANTI, "F10." & CNTDC$(CI%)), 50, 11)
'            Call REPLA(TX$, LOTE$, 62, 16)
            Call REPLA(TX$, FORMATNUM$(USOII, "F12." & CNTDC$(CI%)), 85, 12)
            TRAZCOMPO.List2.AddItem TX$
         End If
       End If
       '
   Next YX&
   
   'Presentar formulario para asignar lotes
   If TRAZCOMPO.List2.ListCount > 0 Then
30   TRAZCOMPO.Show 1
     If TRAZCOMPO.List2.ListCount < 1 Or TRAZCOMPO.Label7 = "FALTAN_LOTES" Then
          TRAZCOMPO.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         'PRIMERO VALIDO QUE TODOS TENGAN LOTE
         For I& = 1 To TRAZCOMPO.List2.ListCount
           LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(I& - 1), 62, 16))
           CI% = CODINT%(Mid$(TRAZCOMPO.List2.List(I& - 1), 1, 16))
           If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote Para el Código: " & CODEXT(CI%))
              GoTo 30 ' PRESENTA NUEVAMENTE EL TRAZCOMPO
           End If
         Next I&

         j& = 0
         For YX& = 1 To ULTREG&(ARCHIVO1$)
          YY$ = REGLEIDO$(ARCHIVO1$, YX&)
          CI% = CVI(Mid$(YY$, 1, 2))
          CAN# = CVD(Mid$(YY$, 19, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If ESTRAZABLE%(CI%) < 1 Then
                 j& = j& + 1
                 Call GRAREG(ARCHIVO1$, YY$, j&)
              End If
            End If
          End If
        Next YX&
        FE% = HOY(HO$)
        '
        For I& = 1 To TRAZCOMPO.List2.ListCount
             'CODIGO DE LIST DE TRAZCOMPO
              CI% = CODINT%(TRIM$(Left$(TRAZCOMPO.List2.List(I& - 1), 16)))
              CANTI = XVALO(Mid$(TRAZCOMPO.List2.List(I& - 1), 50, 11))
              LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(I& - 1), 62, 16))
              USOII = XVALO(Mid$(TRAZCOMPO.List2.List(I& - 1), 85, 12))
              '
              j& = j& + 1
              ZZ$ = ""
              Call REPLA(ZZ$, MKI$(CI%), 1, 2)
              Call REPLA(ZZ$, MKS$(CANTI), 19, 4)   ' VALE A EMITIR
              Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
              Call REPLA(ZZ$, LOTE$, 49, 16)
              Call GRAREG(ARCHIVO1$, ZZ$, j&)
        Next I&
        '
        Call SETULTREG(ARCHIVO1$, j&)
     End If
   End If
   '
   TRAZCOMPO.List2.Clear
 
End Sub


Sub TRANSFEGABAL()
        Dim TOTALRES() As Single
        ReDim TOTALRES(NUMAS%)
        '
        ARM% = 0
        Call PoneEnCeroNull("ORDEFABR", "ARMADOR", REGAF&)
10      OPX% = COMALTER%("Elija la Opcion de Transferencia\\Interna\Armadores")
        If OPX% < 1 Or OPX% > 2 Then Exit Sub
        If OPX% = 2 Then
            Call SELECHO("ARMADORS")
            ARM% = XVALO(VALACT1$("ARMADORS"))
            If ARM% = 0 Then Exit Sub
        End If
        '
        DepOrig% = CONTROL$("", "DEPOCEN")
        DEPODEST% = CONTROL$("", "DEPODEST")
        If ARM% > 0 Then
            TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARM%))
            ARMADOR$ = TRIM$(Mid$(TX$, 3, 64))
            Deposito% = Asc(Mid$(TX$, 71, 1))
            NPROV% = CVI(Mid$(TX$, 68, 2))
            If Deposito% < 1 Then
                Call COMUNI("Imposible Continuar. Falta Deposito Para el Armador Elegido")
                GoTo 10
            End If
            DEPODEST% = CStr(Deposito%)
        End If
        '
        If ECOARCH$("TADEPOSI", MKI$(DepOrig%)) = "" Then
            Call COMUNI("Imposible Continuar. Falta Depósito de Origen")
            GoTo 10
        End If
        '
        If ECOARCH$("TADEPOSI", MKI$(DEPODEST%)) = "" Then
            Call COMUNI("Imposible Continuar. Falta Depósito de Destino")
            GoTo 10
        End If
        '
        VDEF$ = String$(64, 0)
        Call REPLA(VDEF$, "VARIAS", 1, 12)
        Call REPLA(VDEF$, Chr$(DEPODEST%), 14, 1)
        Call REPLA(VDEF$, Chr$(DepOrig%), 13, 1)
        '
        OBX$ = OBJPLA$("SEORFADG", VDEF$)
        If OBX$ = "" Then Exit Sub
        '
        DepOrig% = Asc(Mid$(OBX$, 13, 1))
        DEPODEST% = Asc(Mid$(OBX$, 14, 1))
        '
        For I% = 1 To NOMAS%: TOTALRES(I%) = 0: Next I%
        '
        'CALCULA LA RESERVAS PENDIENTES
        Call ABRECONEXION
        SQLX$ = "select reserva.idsubor, reserva.cod_inte, reserva.cantres, reserva.cantcons, reserva.codiempr from reserva WITH(NOLOCK)"
        Set RESERTMP = FLEXCONN.Execute(SQLX$)
        '
        With RESERTMP
          Do While (.EOF = False)
             OF$ = TRIM$(SAFETEXT$(!IDSUBOR))
             CodEm% = XVALO(!CodiEmpr)
             CONDI$ = "IDSUBOR='" & OF$ & "' AND Codiempr=" & CodEm%
             If ARM% = 0 Then
                ' SE VERIFICA SI LA ORDEN ES VIEJA. ANTES DEL PROCESO DE ETIQUETADO
                Set ETIQUETATMP = FLEXCONN.Execute("SELECT COUNT(*) AS TOTAL FROM ORFAETI WITH(NOLOCK) WHERE " & CONDI$)
                On Error Resume Next
                TOTETI% = XVALO(ETIQUETATMP!Total)
                ETIQUETATMP.Close
                Set ETIQUETATMP = Nothing
                On Error GoTo 0
                If TOTETI% < 1 Then GoTo 13
                '
             End If
             ' OBTIENE EL ARMADOR DE LA OF DE LA RESERVA
             Set ARMADORTMP = FLEXCONN.Execute("SELECT ARMADOR FROM ORDEFABR WITH(NOLOCK) WHERE " & CONDI$)
             On Error Resume Next
             ARMAD% = XVALO(ARMADORTMP!ARMADOR)
             ARMADORTMP.Close
             Set ARMADORTMP = Nothing
             On Error GoTo 0
             '
             If ARM% = ARMAD% Then
                ' ACUMULA EL MATERIAL PENDIENTE DE TRANSFERIR
                CII% = XVALO(!COD_INTE)
                CRES = XVALO(!CANTRES)
                CCON = XVALO(!CANTCONS)
                TOTALRES(CII%) = TOTALRES(CII%) + (CRES - CCON)
             End If
13        .MoveNext
          Loop
       End With
       RESERTMP.Close
       Set RESERTMP = Nothing
           
       ' GUARDA LA RESERVA CALCULADA ANTERIORMENTE EN EL ARCHIVO RESERVA
       JJ& = 0: CONTA% = 0
       REBLA$ = REGBLAN("!RESERVA")
       Call SETULTREG("!RESERVA", 0)
       For I% = 1 To NUMAS%
            If TOTALRES(I%) > 0.005 Then
              y$ = REBLA$
              Call REPLA(y$, MKI$(I%), 1, 2)
              Call REPLA(y$, MKI$(I%), 15, 2)
                Res# = TOTALRES(I%)
              Call REPLA(y$, MKD$(Res#), 19, 8)
              '
              JJ& = JJ& + 1
              Call GRAREG("!RESERVA", y$, JJ&)
            End If
       Next I%
       Call SETULTREG("!RESERVA", JJ&)
       Call CIERRARCH("!RESERVA")
       '
       ' CALCULA LA CANTIDAD QUE DEBE ENTREGARSE TENIENDO EN CUENTA EL STOCK Y EL LOTE INDIVISIBLE
       FIN& = ULTREG&("!RESERVA")
       For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("!RESERVA", U&)
          CI% = CVI(Left$(XX$, 2))
          CANTI = CVD(Mid$(XX$, 19, 8))
          '
          If CANTI < 0 Then CANTI = 0
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then CANTI = CANTI / 1000
          '
          STODEST = STODEPOS(CI%, DEPODEST%)
          '
          STOCENT = STODEPOS(CI%, DepOrig%)
          If STOCENT < 0 Then STOCENT = 0
          NECNETA = CANTI - STODEST
          If NECNETA < 0 Then NECNETA = 0
          '
          'VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
          CANETA# = NECNETA
          VALEMATX = REDON_INDIVIS(CI%, CANETA#, "COMPRAS")
          If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
             VALEMATX = STOCENT
          End If
          '
          If VALEMATX < 0.05 Then VALEMATX = 0
           '
          ZZ$ = REGBLAN$("VALEMATG")
          Call REPLA(ZZ$, MKI$(CI%), 1, 2)
          Call REPLA(ZZ$, MKS$(CANTI), 3, 4)       ' NECESIDAD BRUTA
          Call REPLA(ZZ$, MKS$(STODEST), 7, 4)     ' STOCK DEPOSITO DESTINO
          Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
          Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
          Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
          Call GRAREG("!VALEMATG", ZZ$, U&)
       Next U&
       Call SETULTREG("!VALEMATG", FIN&)
       Call CIERRARCH("!VALEMATG")
       '
25     TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", Chr$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
       VTB% = VENTABU%("VALEMATG" + "/TITUPAN=" + TITUPA$)
       If VTB% < 0 Then Exit Sub
       '
       If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
       '
       If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then Exit Sub
       '
       ' ASIGNA LOTE A LOS ARTICULOS TRAZABLES
       HAYITEMSTRAZA% = 0
       ASIGLOT07.List1.Clear
       FIN& = ULTREG&("!RESERVA")
       For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            XX$ = REGLEIDO$("!VALEMATG", U&)
            CCII% = CVI(Left$(XX$, 2))
            CANTI = CVS(Mid$(XX$, 19, 4))
            If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 And CANTI > 0.05 Then
               Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
               Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
               Call REPLA(TXTX$, FORMATNUM$(CANTI, "F10.1"), 54, 12)
               Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
               ASIGLOT07.List1.AddItem (TXTX$)
               HAYITEMSTRAZA% = 1
            End If
       Next U&
       '
       If HAYITEMSTRAZA% = 1 Then ASIGLOT07.Show 1
       '
       ' GRABACION DE LOS MOVIMIENTOS DE CONSUMO
       CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
       If CMOV$ = "" Then CMOV$ = "TR"
       '
       NRO& = ProNroComprobante(CMOV$)
       '
       Call BLOQARCH("MOVISTO")
       '
       Screen.MousePointer = 11
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       'LOTE$ = ""
       FECHA% = HOY(HOS$)
       DOPRI$ = "TR." + TRIM$(Str$(NRO&))
       'DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
       DOSEC$ = TRIM$(ORFA$) ' NUMERO DE ORDEN DE FABRICACION
        While Len(DOSEC$) < 10
          DOSEC$ = DOSEC$ + " "
        Wend
       '
       NORI% = 0
       'AQUI FALTABA LEER VALEMATG
       For U& = 1 To ULTREG&("!VALEMATG")
          XX$ = REGLEIDO$("!VALEMATG", U&)
          If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
             '
             CI1% = CVI(Mid$(XX$, 1, 2))
             CAN = CVS(Mid$(XX$, 19, 4))
             REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
             '
             CANTI = (-1) * CAN
             SALDOPEND = CAN
             NORFI% = NORI% + 1
             '
'             If ESTRAZABLE%(CI1%) > 0 Then
'                For II& = 0 To ASIGLOT07.List1.ListCount
'                    XX$ = ASIGLOT07.List1.List(II&)
'                    If XVALO(Mid$(XX$, 80, 5)) = U& Then
'                     If CODINT%(TRIM$(Mid$(XX$, 1, 16))) = CI1% Then
'                       CANTLOT$ = Mid$(XX$, 85)
'                       GoTo 35
'                     End If
'                    End If
'                Next II&
'                '
'                ' MUEVE EL STOCK POR CADA LOTE
'35              NN% = InStr(CANTLOT$, ";")
'                If NN% > 0 Then
'                    LOTE$ = Mid$(CANTLOT$, 1, NN% - 1)
'                    If TRIM$(LOTE$) <> "" Then
'                         CANTI = SaldoLote(CI1%, LOTE$, DepOrig%)
'                         If CANTI > SALDOPEND Then CANTI = SALDOPEND
'                         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, (-1) * CANTI)
'                         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
'                         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
'                         CANTLOT$ = Mid$(CANTLOT$, NN% + 1)
'                         SALDOPEND = SALDOPEND - CANTI
'                         GoTo 35
'                    End If
'                End If
'                GoTo 40
'             End If
             '
             LOTE$ = TRIM$(Mid$(XX$, 49, 16)) ' INICIALIZAMOS POR SI NO ENTRA NO QUEDE ANTERIOR
             Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
             '
             REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
             '
             CANTI = CAN
             Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
             '
          End If
40     Next U&
       '
       Call LIBARCH("MOVISTO")
       
       Call BAJALDISCO
       '
       'IMPRIME EL VALE DE TRANSFERENCIA
       '
       If ARM% < 1 Then
'60         FORIPRE$ = CONTROL$("", "FORVALMA")
           FORIPRE$ = ""
           If FORIPRE$ <> "" Then
              '
              FECHDOC$ = HOS$
              NUMEDOC$ = TRIM$(Str$(NRO&))
              While Len(NUMEDOC$) < 6
                 NUMEDOC$ = "0" + NUMEDOC$
              Wend
              TIPODOC$ = "Vale de Materiales"
              '
              CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
              CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(Str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%)))
              CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUORFAB$)
              CAPARDOC% = 3
              '
              CODTABU1$(1) = "COD-MAT"
              CODTABU1$(2) = "DES-MAT"
              CODTABU1$(3) = "UNIMED"
              CODTABU1$(4) = "CANTIDAD"
              CODTABU1$(5) = "CAKILOS"
              CODTABU1$(6) = "REF-MAT1"
              CACOLTAB1% = 6
              '
              DESTIDOC% = 0
              CAITAB1% = 0
              CAITCRU% = 0
              '
              For U& = 1 To ULTREG&("!VALEMATG")
                 XX$ = REGLEIDO$("!VALEMATG", U&)
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
                    TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
                    '
                 End If
              Next U&
              '
              If CAITAB1% < 1 Then
                 Screen.MousePointer = 1
                 Exit Sub
              End If
              '
              Call PRINTDOC("FORVALMA")   ' Vale de Materiales
              Call BAJALDISCO
              '
           End If
       Else
            NPROVE% = NPROV%
            Call OFABRI07.GeneraRemiConsig(NUORFAB$, NPROVE%, NUMREMI$)
            Call ACTUREGISTRO("MOVISTO", "DOSECLAR", "DOPRILAR='TR." & NRO& & "' AND CANTSALID>0", NUMREMI$, REGAF&)
            Call ACTUREGISTRO("MOVISTO", "DOSECCOR", "DOPRILAR='TR." & NRO& & "' AND CANTSALID>0", NUMREMI$, REGAF&)
            Call ACTUREGISTRO("MOVISTO", "Nume_Prov", "DOPRILAR='TR." & NRO& & "' AND CANTSALID>0", NPROVE%, REGAF&)
       End If
       '
       Call COMUNI("Transferencia Realizada")
       Screen.MousePointer = 1
            
End Sub
Sub ORFNUE(REPIM%)                           ' NUEVA ORDEN DE FABRICACION
    '
    ' REPIM%=1 ==> nueva orden de fabricacion
    ' REPIM%=2 ==> repite impresion de O/FAB
    '
    Static YADEPU%
    
    MODOGABAL% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "GABAL") > 0 Then MODOGABAL% = 1
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    Dim CILOTE%(256), CALOTE#(256), CALOTE1#(256), PUNMAT$(256), PUNOPE$(256)
    Dim CISUMI%(512), CASUMI#(512)
    '
'    EJEMPLO DE APLIXYZ$ QUE VIENE DE PEDIDO: "ORFNUE|719|1|2|2479|€|720|2|2|2479|€|721|3|2|2479|"
50  POS1% = InStr(APLIXYZ$, "€")
    If POS1% > 0 Then
       RESTO$ = Mid$(APLIXYZ$, POS1% + 1)
       If Left$(RESTO$, 6) <> "ORFNUE" Then RESTO$ = "ORFNUE" & RESTO$
       If Right(RESTO$, 1) = "€" Then RESTO$ = Left$(RESTO$, Len(RESTO$) - 1)
       APLIXYZ$ = Left$(APLIXYZ$, POS1% - 1)
    Else
       RESTO$ = ""
    End If
    ORFADIRECT% = 0 'validacion si es orden de fabricacion directa desde pedido
    If CONTROL$("ADMIPED", "OFDIRECT") = "SI" And APLIXYZ$ <> "" Then ORFADIRECT% = 1
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
       'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       Dim ORFATEMP As ADODB.Recordset
       Set ORFATEMP = New ADODB.Recordset
       ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
       '
       On Error Resume Next
       With ORFATEMP
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No Hay O/F's Pendientes\que Puedan Refrescarse.")
           Screen.MousePointer = 1
           On Error GoTo 0
           Exit Sub
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While Not .EOF
           If Val(Mid$(!IDSUBOR, 11)) = 1 Then
             ZZ$ = REGBLAN$("ORFANUE")
             REFXX$ = Left$(TRIM$(Left$(CODEXT$(!COD_INTE), 15)) + " - " + TRIM$(DESCRIT0$(!COD_INTE)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
             Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
             Call REPLA(ZZ$, REFXX$, 5, 64)
             JJ& = JJ& + 1
             Call GRAREG("!ORFANUE", ZZ$, JJ&)
           End If
         .MoveNext
         Loop
       End With
       Set ORFATEMP = Nothing
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
       SQLX$ = "SELECT COD_INTE, CanProy, FechGen, FechLan, FechEnSol, SecciPro, DESTINO, StatuOrf FROM ORDEFABR WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE NumeOrFa = " & NRO& & " "
       SQLX$ = SQLX$ + "order by IdSubOr ASC"
       'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set ORFATEMP = New ADODB.Recordset
'       ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
       Set ORFATEMP = READSET(SQLX$)
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
              CIOF% = !COD_INTE
              LOTE# = !CanProy
              Status% = !StatuOrf
            End If
            If !COD_INTE = CIOF% Then CAN# = CAN# + !CanProy
            FGEN = CVINT(!FechGen)
            FLAN% = CVINT(!FechLan)
            FENTRE% = CVINT(!FechEnSol)
              If FENTRE% < HOY(HOS$) Then FENTRE% = HOY(HOS$)
            SPX$ = MKS$(!SecciPro)
            SECPROD = CVS(SPX$)
            DESTI$ = !DESTINO
            If !StatuOrf > 1 Then
              Call MENSERR(24, "Imposible Regenerar.\O/F Cumplida, Suspendida o Anulada.")
              GoTo 100
            End If
         .MoveNext
         Loop
       End With
       Set ORFATEMP = Nothing
       CI% = CIOF%: GoTo 104
       '
1005       On Error GoTo 0
       Call MENSERR(24, "Orden no Encontrada")
       GoTo 100
    End If
    '
    GoSub ARMAPRIMERA
    If ORFADIRECT% = 1 Then X1$ = X0$: GoTo 33
    X1$ = OBJPLA$(OBJEPRI$, X0$)
    If X1$ = "" Then GoTo 999
    '
33  GoSub VERICODI
    If CIOF% < 1 Then GoTo 100
    ' validacion que no tenga marca en proceso de articulos con salidas menores a 5000 e
    ' importes por cantidad salida menores a 5000
    If STATMOVI%(CIOF%) < 0 Then
       Call COMUNI("El Artículo " & CODEXT$(CIOF%) & " Tiene Marca de Baja Para Realizar Ordenes de Fabricación ")
       GoTo 100
    End If
    '
    CIXI% = CIOF%
    SPX$ = String$(4, 0): SPY$ = ""
    '
    If ULTREG&("SECPROD") > 0 Then
      If ORFADIRECT% = 1 Then OBX$ = MKI$(CIXI%) + MKS$(SEPRODUC(CIXI%)): GoTo 1011
101   OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + MKS$(SEPRODUC(CIXI%)))
      
      If OBX$ = "" Then GoTo 999
1011  SPX$ = Mid$(OBX$, 3, 4)
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
    If ORFADIRECT% = 1 Then XX22$ = X1$ + Chr$(TMT%): GoTo 1021
102 XX22$ = OBJPLA$("ORFANUE-2", X1$ + Chr$(TMT%))
1021 X2$ = Left$(XX22$, 256)
    REFF$ = Mid$(XX22$, 258, 30)
    If X2$ = "" Then
'       If APLIXYZ$ <> "" Then
'          Call FINAL("")
'          Exit Sub
'       End If
        'ESTO ES PARA CUANDO SE PASA DESDE PEDIDO GNERACION DE O.F. AUTOMATICAS
        If RESTO$ <> "" Then
            APLIXYZ$ = RESTO$
            GoTo 50
        End If
        If APLIXYZ$ <> "" Then
            Call COMUNI("Generación de O.F. Automática Terminado")
            Call FINAL("")
        End If

       GoTo 100
    End If
    '
    'ASIGNO EL DEPOSITO DE PAÑOL               '
    DEPOPAÑOL% = XVALO(CONTROL$("", "DEPOCEN"))
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
103     XACCE$ = ""
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
    If ORFADIRECT% = 1 Then GoTo 1041
    For II% = 1 To NORDE%
      If CANOPS%(CILOTE%(II%)) < 1 Then
        CMAT% = COMALTER%("Falta Definir Operaciones para\" + TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT0$(CILOTE%(II%))) + "\\Cancelar\Continuar")
        If CMAT% <> 2 Then GoTo 102
      End If
    Next II%
    '
1041 GoSub GERESMAT      ' genera reserva de componentes
    '
       If VACONTROL% < 0 Then
         If REPIM% = 2 Then
            GoTo 100
          Else
            GoTo 102
         End If
       End If
       '
    Call BORRARESERVA(NORFA$)
    '
    If MODOGABAL% <> 1 Then
        Call ASIGNALOTE_RESERVA(NORFA$, OK%) 'trazabilidad de componentes
        If OK% < 0 Then GoTo 1041
    End If
   
    'Call CIERRARCH("ORFANUE")
    Call CIERRARCH("*.*")
    '
    TERENUME$ = ""
    If REPIM% <> 2 Then
       '
       NROPROX& = 1
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
       SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
       'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       Dim ORFABTEMP As ADODB.Recordset
       Set ORFABTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       On Error Resume Next
       With ORFABTEMP
         .MoveFirst
         If Err < 1 Then
           NROPROX& = 1 + !NUMEORFA
         End If
       End With
       On Error GoTo 0
       '
       If NROPROX& > NRO& Then ' SE MODIFICO PARA QUE CUANDO LA NUMERACION SEA MANUAL DE TODOS MODOS VALIDE QUE EL NUMERO INGRESADO NO SEA MENOR AL  MAXIMO INGRESADO.
          TERENUME$ = "FLEXOFT ha Re-Numerado la O/F\como Orden Número '" + TRIM$(Str$(NROPROX&)) + "'."
          NRO& = NROPROX&
       End If
    End If
    '
    If CONTROL$("", "FORMUCON") <> "MULTILIN" Then ' si no esta marcado como multinivel no realiza el
      If ContRepetCompon% > 0 Then GoTo 999 ' CONTROL DE REPETICION DE COMPONENTES
    End If
    '
    GoSub GRAORFNUE
    Call CIERRARCH("ORFANUE")
    GoSub GRAORDEFA
    Call CIERRARCH("ORDEFABR")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ARCON") > 0 And TRIM$(CODFAMIL$(CIOF%)) <> "" Then   'GENERACION DE NUMEROS DE SERIE ARCON
        GoSub GENERA_SERIE_ARCON
    End If
    GoSub GRARESERV
    GoSub DEPUORFEM
    '
    GoSub GRASECOP
    '
 If CONTROL$("TRAZABI", "MODOTRAZ") = "AUTOMA" Then
    GoSub TRAZAUTOM
 End If
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
    'VALIDACION SI TIENE ESTRUCTURA VALIDA QUE SE HAYAN GUARDADO LAS RESERVAS
    If CantRegistros("FORMULAS", "CODICONJ = " & CIOF%, "NOMESS") > 0 Then
       If CantRegistros("RESERVA", " IDSUBOR = '" & NORFA$ & "'", "NOMESS") < 1 Then
          If COMALTER%("Orden de Fabricación Generado Sin Reservas De Materiales,\Desea Genearar Ahora Las Mismas\\Si Reservar\No, Continuar") <> 2 Then
              Call RESERV07.COMMAND3_Click
          End If
       End If
    End If
    'FIN VALIDACION
    '
    'SI EL CODIGO ES TRAZABLE = 2 Y EXISTE
    'EL TRANSE07 VA DIRECTO AL GERADOR DE NUMEROS DE SERIE
    If EXISTE%("TRANSE07.EXE") > 0 And ESTRAZABLE(CIOF%) = 2 Then
        Call CONECRUN("TRANSE07/GENSEOF" + NORFA$, "Generador de Numeros de Serie")
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    If TERENUME$ <> "" Then
       Call COMUNI(TERENUME$)
    End If
    '
'    If APLIXYZ$ <> "" Then
'       Call FINAL("")
'       Exit Sub
'    End If
    'ESTO ES PARA CUANDO SE PASA DESDE PEDIDO GNERACION DE O.F. AUTOMATICAS
    If RESTO$ <> "" Then
        APLIXYZ$ = RESTO$
        GoTo 50
    End If
    If APLIXYZ$ <> "" Then
        Call COMUNI("Generación de O.F. Automática Terminado")
        Call FINAL("")
    End If

    '
    GoTo 100
    '
    '
' ************************************* subrutinas asociadas
    '
FORIPRE:
    FORIPNOR$ = TRIM$(CONTROL$("", "FORMORFA"))
    If FORIPNOR$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPNOR$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPNOR$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Fabricación:\  \'" + RMCC$ + "'.")
            FORIPNOR$ = ""
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
       If APLIXYZ$ <> "" Then GoTo 402
401        MODORFA% = COMALTER%("Puede Generar Ordenes de Fabricación\en Forma Directa, sobre la Base\del Cálculo MRP o de una O/Venta\\Directa\MRP\O/Venta\Cancelar")
402        If MODORFA% = 1 Then
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
    'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORFABTEMP1 As ADODB.Recordset
    Set ORFABTEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    On Error Resume Next
    With ORFABTEMP1
      .MoveFirst
      If Err < 1 Then
        NRO& = 1 + !NUMEORFA
      End If
    End With
   On Error GoTo 0
   Set ORFABTEMP1 = Nothing

    '
    X0$ = REGBLAN$("ORFANUE")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 69, 2)
    Call REPLA(X0$, MKI$(HO%), 109, 2)
    Call REPLA(X0$, MKI$(HO%), 113, 2)
    '
    If APLIXYZ$ <> "" Then
       CODICAN6$ = Mid$(APLIXYZ$, 8)
       PPXX% = InStr(CODICAN6$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(CODICAN$)
       CODI6$ = Left$(CODICAN6$, PPXX% - 1)
       CANT6 = XVALO(Mid$(CODICAN6$, PPXX% + 1))
       CODICAN6$ = Mid$(CODICAN6$, PPXX% + 1)
       PPXX% = InStr(CODICAN6$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(CODICAN$)
       NC% = XVALO(Mid$(CODICAN6$, PPXX% + 1))
       CODICAN6$ = Mid$(CODICAN6$, PPXX% + 1)
       PPXX% = InStr(CODICAN6$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(CODICAN$)
       NUMPEDI% = XVALO(Mid$(CODICAN6$, PPXX% + 1))
       CODICAN6$ = Mid$(CODICAN6$, PPXX% + 1)
       PPXX% = InStr(CODICAN6$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(CODICAN$)
       ORCOMCLI$ = (Mid$(CODICAN6$, PPXX% + 1))
       Call REPLA(X0$, MKI$(CODINT%(CODI6$)), 71, 2)
    End If
    '
    
    Return
    '
'**************************************************************************
VERICODI:                                                             ' esta OK
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
710     Call REPLA(YACCE$, TALLE$(CIOF%), 3, 12)
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
       If ORFADIRECT% = 1 Then XACCE$ = X0$: GoTo 711
       XACCE$ = OBJPLA$("BORENACCE", YACCE$)
    End If
    '
711 If XACCE$ <> "" Then
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
    REFCOMPXX$ = ""
    TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODEXT$(CIOF%)) + " - " + DESCRIT0$(CIOF%)
    '
    VTB% = 0
    If REPIM% <> 2 Then
       VTB% = VENTABU%("BCOMPLEM/TITUPAN=Orden de Complementos:" + TITU$)
    End If
    '
    If VTB% >= 0 Then
       If HAYCOMP% > 0 Then
           j& = 0
           For U& = 1 To ULTREG&("BCOMPLEM")
             YCC$ = REGLEIDO$("BCOMPLEM", U&)
             If CVI(Left$(YCC$, 2)) <> CIOF% Then
                j& = j& + 1
                Call GRAREG("BCOMPLEM", YCC$, j&)
             End If
           Next U&
         Else
           j& = ULTREG&("BCOMPLEM")
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
          j& = j& + 1
          Call GRAREG("BCOMPLEM", YCC$, j&)
          '
          If CACOM > 0 Then
             If REFCOMPXX$ <> "" Then REFCOMPXX$ = REFCOMPXX$ + " - "
             REFCOMPXX$ = REFCOMPXX$ + TRIM$(Mid$(YCC$, 3, 14))
          End If
       Next U&
       Call SETULTREG("BCOMPLEM", j&)
    End If
    Return
    '
'**************************************************************************
AJUCANORF:                                                            ' esta OK
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
'''  ESTO MODIFICADO POR GABAL DEBAJO SE MODIFICO EL CODIGO
'''    If LINDIVIS(CIOF%) >= 1 Then
'''       CANORF1 = (1 + Int((CANORF - 1) / LINDIVIS(CIOF%))) * LINDIVIS(CIOF%)
'''       If Abs(CANORF - CANORF1) > 0.5 Then
'''          CANORF = CANORF1
'''          Call MENSERR(24, "Cantidad Ajustada a Múltiplo\de Lote Indivisible.")
'''       End If
'''    End If
    '
    XANTI = LINDIVIS(CIOF%)
    ' GABAL - COMPLETA CANTIDAD CON EL LOTE INDIVISIBLE DE VENTA
    If MODOGABAL% > 0 Then
        XANTI = LINDIVEN(CIOF%)
        If XANTI < 0.05 Then
            Call COMUNI("Imposible Continuar.\Falta Lote Indivisible de Ventas del Articulo")
            GoTo 100
        End If
    End If
    If XANTI >= 1 Then
       CANORF1 = (1 + Int((CANORF - 1) / XANTI)) * XANTI
       If Abs(CANORF - CANORF1) > 0.5 Then
          CANORF = CANORF1
          Call MENSERR(24, "Cantidad Ajustada a Múltiplo\de Lote Indivisible.")
       End If
    End If
    '
    If LOSTA < CANORF Then
      XLMX$ = MKI$(CI%) + MKS$(CANORF)
      If ORFADIRECT% = 1 Then XLMY$ = XLMX$: GoTo 712
      XLMY$ = OBJPLA$("CARLOREPOMA", XLMX$)
      If XLMY$ <> "" Then
712     LOSTA = CVS(Mid$(XLMY$, 3, 4))
        If LOSTA < 1 Then LOSTA = 1
        'Master.FindFirst ("CODINT = " & CIOF%)
        SQLX$ = "UPDATE MASTER SET LOREPOS =" & LOSTA & ""
        SQLX$ = SQLX$ + " WHERE CODINT = " & CIOF% & ""
         FLEXCONN.Execute (FILTSQL$(SQLX$))
        'If Master.NoMatch = False Then
          'Master.Edit
          'Master!LOREPOS = LOSTA
          'Master.Update
        'End If
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
    If APLIXYZ$ <> "" Then
       CANORF = CANT6
       DESTI$ = "VENTA"
    End If
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
    
    ' GABAL - VALIDACION QUE CANTIDAD SEA MULTIPLO DE LOTE INDIVISIBLE DE VENTA
    If MODOGABAL% > 0 Then
      CAN# = CVD(Mid$(X2$, 89, 8))
      CANLINDIV# = ROUND#(REDON_INDIVIS(CIOF%, CAN#, "VENTA"))
      If CANLINDIV# <> CAN# Then
           Call COMUNI("Cantidad del Código: '" & CODEXT$(CIOF%) & "'\Ajustada Según Lote Indivisible")
           Call REPLA(X1$, MKD$(CANLINDIV#), 89, 8)
           Return
      End If
    End If

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
    SECPROD = CVS(SPX$)
    FCIE = 0
    '
    x$ = REGBLAN$("INDIORF")
    Call REPLA(x$, MKS$(NRO&), 1, 4)
    Call REPLA(x$, FECHATEX$(FGEN), 5, 8)
    If Status% > 0 Then Call REPLA(x$, STS$, 13, 3)
    If PRIO% > 0 Then Call REPLA(x$, PRS$, 16, 3)
    Call REPLA(x$, ECOARCH$("SECPROD", MKS$(SECPROD)), 21, 20)
    If FCIE <> 0 Then Call REPLA(x$, FECHATEX$(FCIE), 43, 8)
    '
    TINDIORF$ = x$      ' prepara registro para INDIORF.DAT
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
                For j& = 1 To FIN&
                  x$ = REGLEIDO$("ARTIPRO", j&)
                    If CVI(Left$(x$, 2)) = CIK%(II%) Then
                      If CVI(Mid$(x$, 3, 2)) > 0 Then
                        PAD%(II%) = 0
                        GoTo 29
                      End If
                    End If
                Next j&
            End If
29          Next II%
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
       NORFAXX$ = NBRS$ + TRIM$(Str$(II%))
       Call ANULARES(NORFAXX$)
       If ORFADIRECT% = 1 Then VACONTROL% = 0
       MODOCOM% = 1: If InStr(1, EMPREAC$, "GABAL") > o Then MODOCOM% = 4
       Call CARCOMPO(CODICONJ$, NORFAXX$, CACONJU#, VACONTROL%, MODOCOM%)
       If VACONTROL% < 0 Then Return         ' cancelacion
       '
       HAYFALTA% = 0
       '
 ARCHIVOX1$ = "MATEROF"
' EPAC$ = TRIM$(UCase$(EMPREAC$))
' If InStr(EPAC$, "MEDITEA") > 0 Then ARCHIVOX1$ = "MATEROF9"
       '
       For U& = 1 To ULTREG&("!" & ARCHIVOX1$)
         ZZ$ = REGLEIDO$("!" & ARCHIVOX1$, U&)
         If CVS(Mid$(ZZ$, 25, 4)) > 0.05 Then HAYFALTA% = 1
       Next U&
       If HAYFALTA% > 0 Then
         ' esto es para impedir generar O/F´s con faltantes - ACM 15/12/06
         If DERACCE%("GENORFAL/NOMESS", "Generar O/F con Faltantes de Materiales") < 2 Then
            Call MENSERR(24, "Hay Faltantes de Materiales\para esta Orden de Fabricación.\  \Usuario sin Derechos de Generar\Ordenes de Fabricación con Faltantes.\Puede Solicitar Derechos con Clave 'GENORFAL'.\  \La Transacción se ha Cancelado.")
            VACONTROL% = (-1)
            Return
         End If
         '
         If COMALTER%("Hay Faltantes de Materiales\para esta Orden de Fabricación.\\Cancelar\Generar Igual") <> 2 Then
           VACONTROL% = (-1)
           Return
         End If
       End If
       '
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       '
       For U& = 1 To ULTREG&("!" & ARCHIVOX1$)
          ZZ$ = REGLEIDO$("!" & ARCHIVOX1$, U&)
          CIIX% = CVI(Left$(ZZ$, 2))
          CANTX# = CVS(Mid$(ZZ$, 5, 4))
          USOII = CANTX# / CACONJU#
          ESTRAZ% = ESTRAZABLE%(CIIX%)
          If ESTRAZ% <> 1 Then ESTRAZ% = 0
          '
          'If CANTX# >= 0.05 Then
          If Abs(CANTX#) > 0.000005 Then  ' PARA CONSIDERAR COMPONENTES CON CANTIDADES NEGATIVAS
             YY$ = REGBLAN$("RESERVA")
             Call REPLA(YY$, MKI$(CIIX%), 1, 2)
             Call REPLA(YY$, NORFA$, 3, 12)
             Call REPLA(YY$, MKI$(CIIX%), 15, 2)
             Call REPLA(YY$, MKI$(HO%), 17, 2)
             Call REPLA(YY$, MKD$(CANTX#), 19, 8)
             Call REPLA(YY$, MKD$(CANTX#), 35, 8)
             Call REPLA(YY$, MKS$(ESTRAZ%), 61, 4)
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
'      With OrdeFabr
'        .MoveFirst
'        Do While Not .EOF
'          If !NUMEORFA = NRO& Then
'            .Delete
'          End If
'        .MoveNext
'        Loop
'      End With
      ' BORRA LA ORDEN DE FABRICACION
      'Ofabric.Execute "DELETE * FROM ORDEFABR WHERE NumeOrfa = " & NRO
       FLEXCONN.Execute (FILTSQL$("DELETE  FROM ORDEFABR WHERE NumeOrfa = " & NRO & ""))

    End If
    '
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    For II% = 1 To NORDE%
      '
      '
      NORFA$ = NBRS$ + TRIM$(Str$(II%))
      DESCRI$ = TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT$(CILOTE%(II%)))
      '
      '\\\OT-08-0631-OD-12/08/08///
      '
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, "Generación de de O/Fabricación", 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      TTXYZ$ = TTXYZ$ + Chr$(10) + Chr$(13)
      TTXYZ1$ = Space$(76)
      CA12$ = FORMATNUM(CALOTE#(II%), "F12.1")
      Call REPLA(TTXYZ1$, "     - Cantidad =  " & CA12$, 11, 56)
      TTXYZ1$ = TTXYZ1$ + Chr$(10) + Chr$(13)
      SEPARA$ = String$(88, "=") + Chr$(10) + Chr$(13)
      ANOTA$ = TTXYZ$ + TTXYZ1$ + SEPARA$
      '
      '\\\OT-08-0631-OD-FIN///
      '
      SQLX$ = " SELECT * FROM OrdeFabr WHERE NUMEORFA  < 1"
      Dim ORDEFFFABR As ADODB.Recordset
      Set ORDEFFFABR = New ADODB.Recordset
      'ORDEFFFABR.Open SQLX$, FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
25        On Error Resume Next
      ORDEFFFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
      '
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      '
      With ORDEFFFABR
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
          !NUMEORFA = NRO&
          !Referen = AJUSTI$(DESCRI$, 64)
          !IDSUBOR = Left$(NORFA$, 12)
          !FechGen = Now
          !COD_INTE = CILOTE%(II%)
          !COD_EXTE = Left$(CODEXT$(CILOTE%(II%)), 16)
          !DESCRIP = Left$(DESCRIT0$(CILOTE%(II%)), 64)
          !Unidad = Left$(UNIMED$(CILOTE%(II%)), 4)
          !CanProy = CALOTE#(II%)
          If Len(DESTI$) > 8 Then
          DESTI$ = TRIM$(Left(DESTI$, 8))
          End If
          !DESTINO = DESTI$
          !FechEnSol = CVDAT(FENTRE%)
          !Priorid = PRIO%
          !StatuOrf = Status%
          !FechLan = CVDAT(FLAN%)
          !SecciPro = SECPROD
          !Fechcierr = CVDAT(0)
          !CanApro = 0
          !CanRech = 0
          !IdenRetra = ""
          !FechProyEntre = CVDAT(0)
          !FechTempEntre = CVDAT(0)
          !FechRecal = CVDAT(0)
          !Observa = ""
          '\\\OT-08-0631-OD-12/08/08///
          !AnotaOrf = ANOTA$
          '\\\OT-08-0631-OD-FIN///
          !NROPED = XVALO(NUMPEDI%)
          '
        .Update
      End With
      ORDEFFFABR.Close
      Set ORDEFFFABR = Nothing
      '
      CONDI$ = "IdSubOr = '" & NORFA$ & "'"
      Call ACTUREGISTRO("OrdeFabr", "Nume_Cli", CONDI$, NC%, REG&, "NOMESS")
      Call ACTUREGISTRO("OrdeFabr", "Origen", CONDI$, Str(NUMPEDI%), REG&, "NOMESS")
      Call ACTUREGISTRO("OrdeFabr", "Nume_Ocom", CONDI$, ORCOMCLI$, REG&, "NOMESS")
      Call ACTUREGISTRO("OrdeFabr", "Id_Orfa", CONDI$, REFF$, REG&, "NOMESS")
      '
      CANS$ = " - (" + TRIM$(Str$(CALOTE#(II%))) + ")"
      DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
      ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
      '
      Call FILTERPUTRECORD("ECORDEF", 1, NORFA$, ECORDEF$)
      If Status% = 1 Then
        Call FILTERPUTRECORD("ECORDEP", 1, NORFA$, ECORDEF$)
      End If
      '
    Next II%
Return
    '
GRARESERV:
    Call ABRESERVA
    '
    If REPIM% = 2 Then
      NORFX$ = TRIM$(NORFA$)
      SQLX$ = "DELETE FROM RESERVA WHERE IdSubOr = '" & NORFX$ & "'"
       FLEXCONN.Execute (FILTSQL$(SQLX$))
    End If
    '
    NORIT% = 0
    For U& = 1 To ULTREG&("!RESERVA")
      YY$ = REGLEIDO$("!RESERVA", U&)
      NORFX$ = TRIM$(Left$(NORFA$, 9) + Mid$(YY$, 12, 3))
      Call GRAREG("!RESERVA", YY$, U&)
      '
      CIRF% = CVI(Left$(YY$, 2))
      CANTRX# = CVD(Mid$(YY$, 19, 8))
      USOIX = CVS(Mid$(YY$, 75, 4))
      IDLOT$ = TRIM$(Mid$(YY$, 43, 16))
      DEPOX% = Asc(Mid$(YY$, 59, 1))
      '
      SQLX$ = " SELECT * FROM Reserva WHERE COD_INTE < 1 "
      Dim RESERVVVA As ADODB.Recordset
      Set RESERVVVA = New ADODB.Recordset
      RESERVVVA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      'With Reserva
      With RESERVVVA
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !COD_INTE = CIRF%
        !COD_EXTE = Left$(CODEXT$(CIRF%), 16)
        !DESCRIP = Left$(DESCRIT0$(CIRF%), 64)
        !IDSUBOR = Left$(NORFX$, 12)
        !FechRes = Now
        !CANTRES = CANTRX#
        !CANTCONS = 0
        !Saldo_Entre = CANTRX#
        !IDENLOTE = Left$(IDLOT$, 16)
        !Deposito = DEPOX%
        !USOUNIT = USOIX
           NORIT% = NORIT% + 1
        !NUORITEM = NORIT%
        !NUME_SOLINT = 0 ' INICIALIZACION DE CAMPO
        .Update
      End With
    Next U&
    Set RESERVVVA = Nothing
Return

    '
TRAZAUTOM:
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    JRES& = 0    ' para re-generar las reservas por O/F
    '
    For II% = 1 To NORDE%
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       Call ANULARES(NORFA$)
       '
       PUNMAT$(II%) = ""
       TRAZORF07.Label7 = "AUTO"    ' PARA ACTIVAR TRAZADO AUTOMÁTICO POR FIFO
       TRAZORF07.Label32 = NORFA$
       TRAZORF07.Show 1
       '
       SQLY$ = "SELECT COD_INTE, CantRes, CANTCONS, IdenLote, UsoUnit  FROM RESERVA WITH(NOLOCK) "
       SQLY$ = SQLY$ + " WHERE IdSubOr = '" & NORFA$ & "' "
       'Set RESETEMP = Stocks.OpenRecordset(SQLY$, dbOpenDynaset)
'       Dim RESETEMP As ADODB.Recordset
'       Set RESETEMP = New ADODB.Recordset
'       RESETEMP.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
       Set RESETEMP = READSET(SQLY$)
       With RESETEMP
          On Error Resume Next
15            .MoveFirst
          If Err < 1 Then
             Do While Not .EOF
                YY$ = REGBLAN$("RESERVA")
                   CIIX% = !COD_INTE
                Call REPLA(YY$, MKI$(CIIX%), 1, 2)
                Call REPLA(YY$, NORFA$, 3, 12)
                Call REPLA(YY$, MKI$(CIIX%), 15, 2)
                Call REPLA(YY$, MKI$(HO%), 17, 2)
                   CANRES# = !CANTRES
                   CANCON# = !CANTCONS
                   SALRES# = CANRES# - CANCON: If SALRES# < 0 Then SALRES# = 0
                Call REPLA(YY$, MKD$(CANRES#), 19, 8)
                Call REPLA(YY$, MKD$(CANCON#), 27, 8)
                Call REPLA(YY$, MKD$(SALRES#), 35, 8)
                   LOTEX$ = !IDENLOTE
                Call REPLA(YY$, LOTEX$, 43, 16)
                   USOII = !USOUNIT
                Call REPLA(YY$, MKS$(USOII), 75, 4)
                '
                JRES& = JRES& + 1
                Call GRAREG("!RESERVA", YY$, JRES&)
                PUNMAT$(II%) = PUNMAT$(II%) + MKS$(JRES&)
             .MoveNext
             Loop
          End If
       End With
       RESETEMP.Close
       Set RESETEMP = Nothing
    Next II%
    '
    Call SETULTREG("!RESERVA", JRES&)
    TRAZORF.Label7 = ""
    '
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
          x$ = REGLEIDO$("ORFAEMI", REOFEM&)
          CAMENO# = CVD(Mid$(x$, 49, 8))
          If CAMENO# <= CALOTEX# Then
              CALOTEX# = CALOTEX# - CAMENO#
              CAMENO# = 0
            Else
              CAMENO# = CAMENO# - CALOTEX#
              CALOTEX# = 0
          End If
          Call REPLA(x$, MKD$(CAMENO#), 49, 8)
          Call REPLA(x$, FORMATNUM$(CAMENO#, "F10.1"), 29, 10)
          Call GRAREG("ORFAEMI", x$, REOFEM&)
          If CALOTEX# < 0.05 Then Exit For
        Next KKU&
      End If
    Next II%
    '
    JK& = 0: FIN& = ULTREG&("ORFAEMI")
    For REOFEM& = 1 To FIN&
      x$ = REGLEIDO$("ORFAEMI", REOFEM&)
      If CVD(Mid$(x$, 49, 8)) >= 0.05 Then
        JK& = JK& + 1
        Call GRAREG("ORFAEMI", x$, JK&)
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
      CAPROYE# = CALOTE#(II%)
      '
      If REPIM% = 2 Then
'        With OperFab
'1221      .FindFirst "IdSubOr = '" & NORFA$ & "'"
'          If .NoMatch = False Then
'            .Delete
'            GoTo 1221
'          End If
'        End With
        'Ofabric.Execute "DELETE * FROM OPERFAB WHERE IdSubOr = '" & NORFA$ & "'"
        SQLX$ = "DELETE FROM OPERFAB "
        SQLX$ = SQLX$ + "WHERE IdSubOr = '" & NORFA$ & "'"
         FLEXCONN.Execute (FILTSQL$(SQLX$))
      End If
      '
      SQLX$ = "SELECT * FROM SecOper WITH(NOLOCK) "
      SQLX$ = SQLX$ + "WHERE CodIConj=" & CIJK% & " "
      SQLX$ = SQLX$ + " AND StatOper = 1 "
      SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
      'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
'      Dim SECOTEMP As ADODB.Recordset
'      Set SECOTEMP = New ADODB.Recordset
'      SECOTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      Set SECOTEMP = READSET(SQLX$)
      On Error Resume Next
      SECOTEMP.MoveFirst
      If Err < 1 Then
        With SECOTEMP
          Do While Not .EOF
            NO% = !NUMEOPER
            '
            CAREALI# = 0
            If REPIM% = 2 Then
              CAREALI# = CANREALI(NORFA$, NO%)
            End If
            CAPENDE# = CAPROYE# - CAREALI#
            If CAPENDE# < 0 Then CAPENDE# = 0
            '
            ATO% = 0
            'AQUI CORREGI POR QUE SI AltOper ERA NULL DABA ERROR
            If Not IsNull(!AltOper) Then ATO% = !AltOper
            If ATO% < 1 Then   ' SOLO OPERACIONES BASICAS
              DOPE$ = SAFETEXT$(!DESCOPER)
              EQ1$ = SAFETEXT$(!CodMaq0) 'PARA EVITAR ERROR 91 NULL
              'EQAUX$ = AJUSTI$(!CodMaq1, 6) + AJUSTI$(!CodMaq2, 6) + AJUSTI$(!CodMaq1, 6)
              EQAUX$ = ""
              If Not IsNull(!CodMaq1) Then EQAUX$ = AJUSTI$(!CodMaq1, 6)
              If Not IsNull(!CodMaq2) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq2, 6)
              If Not IsNull(!CodMaq1) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq1, 6)
              LOTESTA = !LOTESTAN: If LOTESTA < 1 Then LOTESTA = 1
              CALOTES = 1
              CAOPS = !CantiOps
                If CAOPS < 0.01 Then CAOPS = 1
              CICLOHE = !PieCiclo
                If CICLOHE < 1 Then CICLOHE = 1
              MINPREP = 60 * !HorsPrep * CALOTES
                If CAREALI# > 0.05 Then MINPREP = 0
                On Error GoTo 0
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
              HOAL% = Int(HORANUM(Time$))
              USUA% = USNBR% Mod 100
              DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
              INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
              PROGBLA$ = Space$(6) + String$(26, 0)
              '
              If Status% <= 1 Then
                SQLX$ = "SELECT * FROM OperFab WHERE COD_INTE < 1 "
                Dim OPERFFFAB As ADODB.Recordset
                Set OPERFFFAB = New ADODB.Recordset
                OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                'With OperFab
                With OPERFFFAB
                  .AddNew   ' YA CONVERTIDO
                  On Error Resume Next
                  !CodiEmpr = CodiEmp%
                  On Error GoTo 0
                    !IDSUBOR = Left$(NORFA$, 12)
                    !COD_INTE = CIJK%
                    !COD_EXTE = Left$(CODEXT$(CIJK%), 16)
                    !DESCRIP = Left$(DESCRIT0$(CIJK%), 64)
                    !NUMEOPER = NO%
                    !DESCOPER = Left$(DOPE$, 48)
                    !MAQUINA = Left$(EQ1$, 6)
                    !DescMaq = Left$(ECOARCH("PADMAQ", EQ1$), 64)
                    !EquiAux = Left$(EQAUX$, 18)
                    !CantiOps = CAOPS
                    !CANTPROG = CAPROYE#
                    !CANTREAL = CAREALI#
                    !CANTPEND = CAPENDE#
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
                Set OPERFFFAB = Nothing
              End If
            End If
          .MoveNext
          Loop
        End With
        Set SECOTEMP = Nothing
      End If
      On Error GoTo 0
    Next II%
    '
    'OperFab.Close
    Return
    '
PRINORFAB:
    Call LIMPIATETAB
237     FORIPRE$ = FORIPNOR$ '***///****'FORIPRENOR$
    If CVS(SPX$) >= 1 Then
      TTXYZ$ = FILTERGETRECORD$("SECPROD", 1, SPX$)
      FORIPALTER$ = TRIM$(Mid$(TTXYZ$, 83, 8))
      If FORIPALTER$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPALTER$ + ".FRM") > 0 Then
          FORIPRE$ = FORIPALTER$
        End If
      End If
    End If
    If TRIM$(FORIPRE$) = "" Then Return
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
    modoalear% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ALEAR") > 0 Then modoalear% = 1 ' para ajustar interlinea de operaciones
    '
    If CONTROL("ORDEFABR", "ORDENA") = "UBICA" Then ' si esta seleccionado orden por ubicacion
        Screen.MousePointer = 11
        FIN& = ULTREG&("!Reserva")
        For JJ& = 1 To FIN&
          Call TRACE(20, JJ&, FIN&)
          x$ = REGLEIDO$("!Reserva", JJ&)
          CI% = CVI(Left$(x$, 2))
          UBIC$ = UBISTO$(CI%)
          Call REPLA(x$, UBIC$, 65, 10)
          Call GRAREG("!Reserva", x$, JJ&)
        Next JJ&
        Call CIERRARCH("!Reserva")
        Call TRACENTRAL("RESERVA", "")
        Call FILESORT("RESERVA", "RESERVA", 14, 14)
        Call FILESORT("RESERVA", "RESERVA", 13, 13, "DESC")
        Call TRALOCAL("RESERVA", "")
        Screen.MousePointer = 1
        '
        PUNMAT$(1) = ""
        For JJ& = 1 To FIN&
           PUNMAT$(1) = PUNMAT$(1) + MKS$(JJ&)
        Next JJ&
        '
    End If
    MARCA1% = 1
    ITNUMERO% = 0
    ' IMPRESION DEL CODIGO DE BARRAS PARA PICKING DEL REMITO EN CONSIGNACION
    CADENA$ = REPLACAR$(NORFA$, "-", "")
    BARRA$ = CODIBARRA$(CADENA$, 2)
    '
    For II% = 1 To NORDE%
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KKINI% = 1
249           KK1% = 1: NH% = 1
250           ATTX% = 0
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
          Call PRINTFORWIN(BARRA$, "REF-ASIE", 0, 0)
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
          'Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0) 'SE MODIFICA LA LLAMADA A ESPECIFI$ X ESPECIFIT$ PARA LA
          Call PRINTFORWIN(ESPECIFIT$(CILOTE%(II%)), "ESP-MAT", 0, 0) 'IMPRESION DE LA ESPECIFICIONES DEL ARTICULOS
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
          If MODOGABAL% > 0 Then
            CANS0$ = CSTRING$(MKD$(LINDIVEN(CILOTE%(II%))), 3, 9, 1, 2)
            Call PRINTFORWIN(CANS0$, "PES-CONJ", 0, 0)
          End If
          '
          Call PRINTFORWIN(TRIM$(REFCOMPXX$), "REF-CONJ", 0, 0)
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
          Call PRINTFORWIN(TRIM$(NUMEROS_DE_SERIE$), "REF-MAT4", 0, 0)
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
            Screen.MousePointer = 11
            FINXX& = Len(PUNMAT$(II%))
            For KK% = KK1% To Len(PUNMAT$(II%)) Step 4
              KKLL& = KK%
              Call TRACE(20, KKLL&, FINXX&)
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
                USOUNII = CVS(Mid$(YY$, 75, 4))
                LOTEMPRI$ = TRIM$(Mid$(YY$, 43, 16))
                '
                If CONTROL("ORDEFABR", "ORDENA") = "UBICA" Then ' si esta seleccionado orden por ubicacion
                    MARCA% = CVS(Mid$(YY$, 61, 4)) ' PARA QUE DEJE UN RENGLON EN BLANCO
                      If MARCA% <> MARCA1% Then ' CUANDO CAMBIAN A  TRAZABLES 1
                        YAC = YAC + INTERLI
                      End If
                    MARCA1% = MARCA%
                End If
                '
                'If CANTX# > 0 Then
                If Abs(CANTX#) > 0.000005 Then ' PARA CONSIDERAR COMPONENTES CON CANTIDADES NEGATIVAS
                  Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                  Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                  Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                    STACTU = STODEPOS(CIXI%, DEPOPAÑOL%): If STACTU < 0 Then STACTU = 0
                  Call PRINTFORWIN(FORMATNUM$(STACTU, "F12.1"), "CANT-DIS", 0, YAC)
                  Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                  Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                  Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                  Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                  Call PRINTFORWIN(LOTEMPRI$, "PAR-MPRI", 0, YAC)
                    CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                  Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                    'USXS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                    USXS$ = CSTRING$(MKS$(USOUNII), 3, 9, 3, 2)
                  Call PRINTFORWIN(USXS$, "USO-MPRI", 0, YAC)
                  Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                  ' NUMERO DE ITEM
                  ITNUMERO% = ITNUMERO% + 1
                  ITN$ = FORMATNUM$(ITNUMERO%, "F3.0")
                  Call PRINTFORWIN(ITN$, "ORD-ITEM", 0, YAC)
                  '
                  STACTU = STODEPOS(CIXI%, DEPOPAÑOL%): If STACTU < 0 Then STACTU = 0
                  '
                  If STACTU < 0 Then STACTU = 0
                  FALTANTE# = CANTX# - STACTU         'OBTIENE LA CANTIDAD FALTANTE DEL ARTICULO
                  If FALTANTE# > 0 Then
                    FALT$ = FORMATNUM$(Abs(FALTANTE#), "F9.1")
                    Call PRINTFORWIN(FALT$, "CANT-FAL", 0, YAC)
                  End If
                  YAC = YAC + INTERLI
                  '
                End If
              End If
            Next KK%
            '
260         YAC = DETDES
            ITNU% = 0
            TMPRE = 0: TMPRO = 0
            If KKINI% < 1 Then KKINI% = 1
            'SE MODIFICO EL 1 POR KKINI% PARA QUE MANTENGA EL VALOR DE LA VUELTA
'            For KK% = 1 To ULTREG&("!OPERFAB")
            For KK% = KKINI% To ULTREG&("!OPERFAB")    'KKINI% To Len(PUNOPE$(II%)) Step 4
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
                If TRIM$(DOPE$) <> "" Then
                   DESOP$ = RTrim(DOPE$) + vbCrLf + vbCrLf + DESOP$
                End If
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
                Call PRINTFORWIN(DESOP$, "OBS-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
                Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
                Call PRINTFORWIN(NUHOPROF$(CILOTE%(II%), NOPE%), "N-HPROF", 0, YAC)
                '
                If modoalear% = 0 Then
                    YAC = YAC + INTERLI
                  Else
                    YAC = YAC + 2 * INTERLI
                End If
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
          'DESDE AQUI CODIGO AGREGADO PARA VALOBADA DEL FRM
          'AQUI AGREGAR LOS PARAMETROS TOMADOS DE LA BASE DE DATOS
          XX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRQ") 'ASIGNO A XX$ LA LECTURA DEL ARCHIVO
          For U1% = 1 To Len(XX$) Step 256 'LO RECORRO (256 ES LA LONGITUD DEL REGISTRO)
              YY$ = Mid$(XX$, U1%, 256)
              CODVARI$ = TRIM$(Left(YY$, 16))
              TABLASE$ = TRIM$(Mid$(YY$, 17, 24))
              If TABLASE$ <> "" Then
                 CAMPOSE$ = TRIM$(Mid$(YY$, 41, 24))
                 If CAMPOSE$ <> "" Then
                   CONDI$ = TRIM$(Mid$(YY$, 65))
31                 TBUS$ = "¿"
                   PPXX1% = InStr(CONDI$, TBUS$)
                   If PPXX1% > 0 Then
                     TBUS$ = "?"
                     PPXX2% = InStr(CONDI$, TBUS$)
                     IZQUI$ = Left$(CONDI$, PPXX1% - 1)
                     DERE$ = Mid$(CONDI$, PPXX2% + 1)
                     AREPLA$ = ""
                     If PPXX2% > PPXX1% + 1 Then
                        AREPLA$ = Mid$(CONDI$, PPXX1% + 1, PPXX2% - PPXX1% - 1)
                     End If
                     If AREPLA$ <> "" Then
                        For KKII% = 1 To CAVALPAR%
                           If CODPAR$(KKII%) = AREPLA$ Then
                              AREPLA$ = VALPAR$(KKII%)
                              GoTo 32
                           End If
                        Next KKII%
                     End If
32                   CONDI$ = IZQUI$ + AREPLA$ + DERE$
                     GoTo 31
                   End If
                   '
                   TTXX$ = SAFETEXT$(VALOBADA(TABLASE$, CAMPOSE$, CONDI$))
                   Call PRINTFORWIN$(TTXX$, CODVARI$, 0, 0)
                   ' ESTO ES PARA RETENER EL VALOR PARA LAS IMAGENES
                   If UCase$(CODVARI$) = "COD-MAT" Then CIACTI% = CODINT%(TTXX$)
                 End If
              End If
          Next U1%
          'FIN DE AGREGADO PARA VALOBADA DEL FRM
          '
          'CODIGO PARA QUE IMPRIMA LAS IMAGENES DISTINTAS AL LOGO
          'ESTO ES TOMADO DE PRINTDOC
          For KKI% = 1 To CAIMINI% + CAIMAGEN%
               CODIMA$ = UCase$(CODIMAGEN$(KKI%))
               If Left$(CODIMA$, 7) = "DOCUGRA" Then
                 ' ESTA PARTE ES PARA INCRUSTAR LOS 4 DOCUMENTOS GRAFICOS
                 If CIACTI% > 0 Then
                   NIMA% = XVALO(Mid$(CODIMA$, 8, 1))
                   If NIMA% <= 3 Then
                     FGRC$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CIACTI%))
                     ARIMA$ = TRIM$(Mid$(FGRC$, 65 + 48 * NIMA%, 48))
                       RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
                       If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
                       If Right$(RUTARGRA$, 1) <> "\" Then RUTARGRA$ = RUTARGRA$ + "\"
                     ARCHIMAGEN$(KKI%) = RUTARGRA$ + ARIMA$
                   End If
                 End If
               End If
               Call PRINTIMAGE(CODIMAGEN$(KKI%), ARCHIMAGEN$(KKI%))
          Next KKI%
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
       NUMEDOC$ = NORFA$
       TIPODOC$ = "Orden de Fabricación"
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
          Call PRINTFORWIN(ESPECIFIT$(CILOTE%(II%)), "ESP-MAT", 0, 0)
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
          Call PRINTFORWIN(TRIM$(REFCOMPXX$), "REF-CONJ", 0, 0)
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
                USOUNII = CVS(Mid$(YY$, 75, 4))
                LOTEMPRI$ = TRIM$(Mid$(YY$, 43, 16))
                '
                Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                Call PRINTFORWIN(LOTEMPRI$, "PAR-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                   'USXS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                   USXS$ = CSTRING$(MKS$(USOUNII), 3, 9, 3, 2)
                Call PRINTFORWIN(USXS$, "USO-MPRI", 0, YAC)
                Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                '
                STACTU = STODEPOS(CIXI%, DEPOPAÑOL%): If STACTU < 0 Then STACTU = 0
                If STACTU < 0 Then STACTU = 0
                FALTANTE# = CANTX# - STACTU         'OBTIENE LA CANTIDAD FALTANTE DEL ARTICULO
                If FALTANTE# > 0 Then
                  FALT$ = FORMATNUM$(Abs(FALTANTE#), "F9.1")
                  Call PRINTFORWIN(FALT$, "CANT-FAL", 0, YAC)
                End If
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
          Call PRINTFORWIN(ESPECIFIT$(CILOTE%(II%)), "ESP-MAT", 0, 0)
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
          Call PRINTFORWIN(TRIM$(REFCOMPXX$), "REF-CONJ", 0, 0)
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
      If COMALTER%("¿Desea Continuar o Repetir\la Impresion de este Documento?\\Continuar\Repetir Impresion") = 2 Then GoTo 237
    End If
    '
    DOCUORIG$ = TRIM$(NORFA$)
    NUMEDOC$ = ""
    '
    ' GABAL - ARMADORES
    If MODOGABAL% > 0 Then
        If COMALTER%("¿Desea Asignar Ahora Un Armador Para Esta Orden?\\No, Continuar\Sí, Asignar") = 2 Then
            Call GRABAARMADOR_GABAL(NORFA$, FENTRE%)
        End If
        If REPIM% <> 2 Then
            Call GRABAETIQUETASGABAL(NORFA$, 1)
            OPYY% = COMALTER%("¿Desea Imprimir las Etiquetas?\\No, Cancelar\Sí, Continuar")
            If OPYY% = 2 Then Call PRINTETIQUETASGABAL(NORFA$, 0)
        End If
    End If
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
GENERA_SERIE_ARCON:

    'OBTENER EL NUMERO DE FAMILIA AL CUAL CORRESPONDE EL P.TERMINADO.
    'CONSULTAR LA TABLA TANUMSE REALIZANDO UN MAX DEL NUMERO DE SERIE WHERE PACIENTE = NUMERO DE FAMILIA
    'SI DEVUELVE 0 PRESENTAR MSGBOX DE INCIALIZACION.
    'VALIDO SI ES TRAZABLE 2
'    If ESTRAZABLE%(CIXI%) < 2 Then
    If TRIM$(CODFAMIL$(CIXI%)) = "" Then
      GoTo 989
    End If
    '
    NUMEROS_DE_SERIE$ = ""
    'OBTENGO EL CODIGO DE FAMILIA
    CODFAMI$ = TRIM$(CODFAMIL$(CIXI%))
    If CODFAMI$ = "" Then GoTo 989
    
    CONDI$ = "PACIENTE =  '" & CODFAMI$ & "'"
    NUEVONUMSERIE = XVALO(VALOBADA("TANUMSE", "MAX(CAST(substring(NUMEROSERIE,1,32) AS INT))", CONDI$, "NOMESS"))
    If NUEVONUMSERIE < 1 Then
5       SERIENUEVO = XVALO(InputBox$("Inicializar Número de Serie", "Ingrese Número"))
        If SERIENUEVO < 1 Then
             OPPX% = COMALTER%(" N.Serie Sin Inicialización :\\Generar con Número 1 \Ingresar Número de Inicialización\Cancelar")
             If OPPX% < 1 Or OPPX% > 2 Then
                Call COMUNI("Esta Orden de Fabriación se Generará sin N.Series")
                'VER A DONDE SE MANDA
             End If
             If OPPX% = 1 Then SERIENUEVO = 1     ' SI GENERA CON NUMERO 1
             If OPPX% = 2 Then GoTo 5                ' SI VUELVE A INGRESAR INICIALIZACION.
        End If
        NUEVONUMSERIE = SERIENUEVO

    Else
        NUEVONUMSERIE = NUEVONUMSERIE + 1
    End If
    '
    NumOrden% = XVALO(VALOBADA("TANUMSE", "MAX(NUMEORD)", CONDI$, "NOMESS"))
    'GRABACION EN TABLA TANUMSE
    
    DESCRI$ = DESCRIT0$(CIXI%)
    Code$ = CODEXT$(CIXI%)
    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE=0"
    
    Dim TANUMTMP As ADODB.Recordset
    Set TANUMTMP = New ADODB.Recordset
    TANUMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For U& = 1 To CANORF
        NUMEROS_DE_SERIE$ = NUMEROS_DE_SERIE$ & SAFETEXT$(NUEVONUMSERIE) & " , "
        With TANUMTMP
            NumOrden% = NumOrden% + 1
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !IDSUBOR = NORFA$
               HOII% = HOY(HOS$)
               FF% = HOII%
            !FechAlta = CVDAT(FF%)
            !COD_INTE = CIXI%
            !COD_EXTE = Code$
            !DESCRIP = DESCRI$
            !MESANIO = ""
            !Ten_Ensa = 0
            !NUMEROSERIE = SAFETEXT$(NUEVONUMSERIE)
            !NUMEORD = NumOrden%
            !DEPONUMSE = (-1)
            !PACIENTE = CODFAMI$
            'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
            !Historial = "Alta N/Serie a Partir de: " & NORFA$
            .Update
            '
        End With
        '
        NUEVONUMSERIE = NUEVONUMSERIE + 1 ' LO INCREMENTO AQUI POR QUE LA PRIMERA VEZ YA TRAE EL SIGUIENTE AL ULTIMO
        '
     Next U&
     TANUMTMP.Close
     Set TANUMTMP = Nothing
     On Error Resume Next
     If InStrRev(NUMEROS_DE_SERIE$, 2) = ", " Then NUMEROS_DE_SERIE$ = Left$(NUMEROS_DE_SERIE$, Len(NUMEROS_DE_SERIE$) - 2)
     On Error GoTo 0
989
    
Return
'-------------
999 Call CIERRARCH("*.*")
    Call FRESHALL
    Screen.MousePointer = 1
'    If APLIXYZ$ <> "" Then
'       Call FINAL("")
'    End If
    'ESTO ES PARA CUANDO SE PASA DESDE PEDIDO GNERACION DE O.F. AUTOMATICAS
    If RESTO$ <> "" Then
        APLIXYZ$ = RESTO$
        GoTo 50
    End If
    If APLIXYZ$ <> "" Then
        Call COMUNI("Generación de O.F. Automática Terminado")
        Call FINAL("")
    End If

End Sub
Sub PRINTETIQUETASGABAL(NORFA$, NUMEETI%)

    CONDI$ = "IdSubOr='" & NORFA$ & "'"
    CAN0# = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$))
    CIOF0% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$))
    FF% = CVINT(VALOBADA("ORDEFABR", "FECHGEN", CONDI$))
    '
    CODIGO$ = CODEXT$(CIOF0%)
    NORFA0$ = NORFA$
    FORIPRE$ = TRIM$(CONTROL$("", "ETIORFA")) 'ETIQUETA STANDARD
    '
    If FORIPRE$ = "" Then
        Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
        Exit Sub ' SI NO HAY NINGUN FRM CONFIGURADO
    End If
    '
30  PIMPRE$ = TRIM$(CONTROL(IDENTER$, "PORTETIQ"))
    If PIMPRE$ = "" Then PIMPRE$ = CONTROL$("*", "PORTETIQ")
    If PIMPRE$ = "" Then Exit Sub ' POR SI NO HAY NINGUNA IMPRESORA CONFIGURADA
    '
    TIPODOC$ = "Etiqueta Orden Fabricacion"
    NUMEDOC$ = NORFA0$
    '
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODIGO$
    CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = NORFA0$
    CODPARAM$(3) = "DESTINO": DOCPARAM$(3) = CODIBARRA_EAN13$(COLOR$(CIOF0%))
    CODPARAM$(4) = "CAN-MPRI": DOCPARAM$(4) = LINDIVEN(CIOF0%)
    CODPARAM$(7) = "FECH-LAN": DOCPARAM$(7) = FECHATEX$(FF%)
    '
    CAPARDOC% = 7
    '
    NOCONFIRM% = 1 ' NO PRESENTA MENSAJE DE REIMPRESION
    CONDI$ = "IDSUBOR='" & NORFA0$ & "'"
    JJ% = 1 ' POR DEFAULT IMPRIMIR DESDE 1 EN ADELANTE. ASUME UNA NUEVA OF
    CANTI_ETIQ_A_GENERAR% = CantRegistros("ORFAETI", CONDI$)
    If NUMEETI% > 0 Then
        CONDI$ = "IDSUBOR='" & NORFA0$ & "' AND NUORDEN=" & NUMEETI%
        If CantRegistros&("ORFAETI", CONDI$) > 0 Then ' VERIFICA QUE EL NUMERO DE ETIQUETA EXISTA
            JJ% = NUMEETI% ' IMPRIME LA ETIQUETA SOLICITADA
            CANTI_ETIQ_A_GENERAR% = NUMEETI%
        Else
            GoTo 80 ' NO ENCONTRO EL NUMERO DE ETIQUETA
        End If
    End If
    '
    PRINTERPORT = ""
    Call SETPRINPORT(PIMPRE$)
    '
    For I% = JJ% To CANTI_ETIQ_A_GENERAR%
        CODPARAM$(5) = "REF-MAT3": DOCPARAM$(5) = "Contador" & I%
            NORFA1$ = REPLACAR$(NORFA0$, "-", "")
            TX$ = I%: While Len(TX$) < 4: TX$ = "0" & TX$: Wend
        CODPARAM$(6) = "REF-CONJ": DOCPARAM$(6) = CODIBARRA$(NORFA1$ & "|" & TX$, 4)
        '
        Call PRINTDOC("@" + FORIPRE$)
'        If I% < CANTI_ETIQ_A_GENERAR% Then
'            Call PRINTDOCNUE("@" + FORIPRE$, "NOPRINT")
'        Else
'            Call PRINTDOCNUE("@" + FORIPRE$, "")
'        End If
    Next I%
    NOCONFIRM% = 0
    '
80  Call SETPRINPORT("RESTORE")
    '
End Sub



