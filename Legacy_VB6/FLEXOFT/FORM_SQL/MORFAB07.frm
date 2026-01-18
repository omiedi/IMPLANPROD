VERSION 5.00
Begin VB.Form MORFAB07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "O/F's - Ordenes de Fabricación"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5010
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "MORFAB07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   35
      Top             =   3780
      Width           =   3690
      Begin VB.PictureBox Picture20 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   550
         ScaleHeight     =   585
         ScaleWidth      =   2595
         TabIndex        =   36
         Top             =   420
         Width           =   2650
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
            Left            =   2035
            Picture         =   "MORFAB07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Listado de Ingresos sin Causa"
            Top             =   0
            Width           =   555
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
            Left            =   0
            Picture         =   "MORFAB07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Control de Cumplimiento de O/F's"
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de O/F's Pendientes"
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
            Left            =   600
            TabIndex        =   37
            Top             =   105
            Width           =   2220
         End
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   24
      Top             =   2415
      Width           =   3690
      Begin VB.Frame Frame16 
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
         TabIndex        =   28
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command71 
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
            TabIndex        =   34
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command70 
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
            TabIndex        =   33
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command69 
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
            TabIndex        =   32
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command44 
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
         Begin VB.CommandButton Command43 
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
            TabIndex        =   30
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command42 
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
            TabIndex        =   29
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture18 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2250
         TabIndex        =   25
         Top             =   420
         Width           =   2310
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
            Picture         =   "MORFAB07.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Transferencia Materiales de Depósito a Producción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignación de Materiales"
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
            TabIndex        =   27
            Top             =   105
            Width           =   1635
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   6840
      Left            =   4095
      ScaleHeight     =   6780
      ScaleWidth      =   1950
      TabIndex        =   8
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "MORFAB07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "MORFAB07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Listado de Informe de Produccion por Fecha"
         Top             =   1410
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   4305
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
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
         Left            =   105
         Picture         =   "MORFAB07.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
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
         Left            =   105
         Picture         =   "MORFAB07.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   705
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "MORFAB07.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2055
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
         Left            =   105
         Picture         =   "MORFAB07.frx":17DA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   3585
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "MORFAB07.frx":1C1C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4620
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Base de Datos de O-F's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture17 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   19
         Top             =   1260
         Width           =   2100
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
            Left            =   0
            Picture         =   "MORFAB07.frx":3B3E
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Anulación de Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de O/Fabricación"
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
            TabIndex        =   20
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   2100
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
            Left            =   0
            Picture         =   "MORFAB07.frx":3E48
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Cierre de Orden de Fabricación - Re-Apertura O/F's con Saldo"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de O/F"
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
            TabIndex        =   18
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   15
         Top             =   420
         Width           =   2100
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
            Picture         =   "MORFAB07.frx":4152
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Generar Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Nueva O/F"
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
            TabIndex        =   16
            Top             =   105
            Width           =   1320
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
         TabIndex        =   1
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   3
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
            TabIndex        =   2
            Top             =   1890
            Width           =   1590
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
      Visible         =   0   'False
      Begin VB.Menu NvaOrd 
         Caption         =   "Nueva Orden"
         Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu modiocomof 
         Caption         =   "Modificar Orden de Compra a O.Fabricación"
      End
   End
End
Attribute VB_Name = "MORFAB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CONCUMORFAB()
  '27/04/07 (OT-07-0172)
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If PERIO$ = "" Then Exit Sub
  DESDEX = CVDAT(DES%)
  HASTAX = CVDAT(HAS%)
  '
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM ORDEFABR "
  SQLX$ = SQLX$ + " WHERE FechGen >= " & DESDEX & " "
  SQLX$ = SQLX$ + " AND FechGen <= " & HASTAX & " "
  SQLX$ = SQLX$ + " AND StatuOrf <> 9 "
  SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC"
  '
  'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim ORDEFABTMP As ADODB.Recordset
  Set ORDEFABTMP = New ADODB.Recordset
  ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  '
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
       Y$ = REGBLAN("COCUMOF")
       FEMIS% = CVINT(!FechGen)
       Call REPLA(Y$, MKI$(FEMIS%), 1, 2)
       Call REPLA(Y$, !IdSubOr, 3, 12)
       Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
       Call REPLA(Y$, MKD$(!CanProy), 17, 8)
       Call REPLA(Y$, !Destino, 25, 8)
       FECHAENT% = CVINT(!FechEnSol)
       Call REPLA(Y$, MKI$(FECHAENT%), 33, 2)
       FECHACIE% = CVINT(!FechCierr)
       Call REPLA(Y$, MKI$(FECHACIE%), 35, 2)
       DIATRAZO% = DIENTRE%(FECHAENT%, FECHACIE%)
       Call REPLA(Y$, MKI$(DIATRAZO%), 37, 2)
       '
       I& = I& + 1
       Call GRAREG("COCUMOF", Y$, I&)
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

Sub REIMPRIETIQUETA(Optional MODO%)
    '
    If MODO% = 1 Then
        Call IMPRE_ETIQUETA(1)
    Else
        ' ///OT-8-1037-12/11/08
        Call SELECHO("MASTER")
        CODEX$ = VALACT1$("MASTER")
        CIXI% = CODINT%(CODEX$)
        If CIXI% < 1 Then Exit Sub
        If ESTRAZABLE%(CIXI%) < 1 Then Exit Sub
        '
        Call ARMALILOT(CIXI%, DEPO%, FAJUI%)
        '
        Call FRESHECHO("!LILOTSE")
        If ULTREG("!LILOTSE") < 1 Then
           Call MENSERR(24, "No Hay Lotes Disponibles\para la Operacion Solicitada.")
           Exit Sub
        End If
        '
        Call SELECHO("!LILOTSE")
        VDEVU$ = VALACT1$("!LILOTSE")
        If TRIM$(VDEVU$) = "" Then Exit Sub
        IDREGIS$ = VDEVU$
        '
        CONDI$ = "IDENLOTE='" & TRIM$(IDREGIS$) & "'"
        DPX% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", CONDI$))
        If DPX% > 0 Then Call PRINSTIKAPRO(CIXI%, IDREGIS$, DPX%)
        ' ///OT-8-1037-FIN
    End If
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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



Private Sub Command14_Click()
    Command14.Enabled = False
    Call COMUNI("No Hay Opciones de Configuración\para esta Aplicación.")
    Command14.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CONECRUN("GENREP07/LISORFA", "Listado de Ordenes de Fabricacion")
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    FEXI% = DIANTE(DIANTE(HOY(HOS$)), 6)
    HOY_7$ = FETEXCOR1$(FEXI%)
    HOY_7TEX$ = FECHATEX$(FEXI%)
    '
    'ALARMA DE AVISO DE INGRESOS DE ORDEN DE FABRICACION SIN CAUSA QUE SUPERARON LOS 7 DIAS
    CONDI$ = " FECHMOV < '" & HOY_7$ & "' AND CAUSA < 1 and Codimovi = 'OF' AND CANTINGRE > 0 "
    CANRE& = CantRegistros&("MOVISTO", CONDI$)
    If CANRE& > 0 Then
       Call FINAL("?CAUSALAR(" & HOY_7$ & ")")
    Else
       Call COMUNI("No Se Encontraron Ingresos Sin Causa\Anteriores a la Fecha: " & HOY_7TEX$)
    End If
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call CONECRUN("TRAZAB07/TRAZORFA", "Asignacion de Materiales a O/F")
    Command4.Enabled = True
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

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
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

Private Sub Command5_Click()
    Command5.Enabled = False
    Call GENORDEFDAT(0) 'NUEVO
    '
    j& = 0
    Screen.MousePointer = 11
    Call HEADERS("ORFASEL", "")
    For U& = 1 To ULTREG&("ORDEFABR")
       X$ = REGLEIDO$("ORDEFABR", U&)
       j& = j& + 1
       Call GRAREG("ORFASEL", X$, j&)
    Next U&
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
99  Command5.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
    '
    ' ///OT-8-1037-12/11/08
    ' RE-IMPRESION DE ETIQUETAS
    OPX% = ALTERNA%("Re-Impresion de Etiquetas (P. Terminados)\Re-Impresion de Etiquetas (M.Primas)\Otros Accesos Directos")
    If OPX% < 0 Or OPX% > 2 Then GoTo 99
    '
    Call REIMPRIETIQUETA(OPX%)
    GoTo 99
    ' ///OT-8-1037-FIN
    '
    TTOP$ = ""
    OPMAX% = 0
    OPTRAZA% = 0
    OPREPA% = 0
    OPSIMU% = 0
    If EXISTE%("TRAZAB07.EXE") > 0 Then
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
    If EXISTE%("SIMULA07.EXE") > O Then
       OPMAX% = OPMAX% + 1
       If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
       OPSIMU% = OPMAX%
       TTOP$ = TTOP$ + "Simulación de Restricciones de Stock"
    End If
    ' '\\\OT-07-0386-FIN///
    '
    If TTOP$ <> "" Then
       TTOP$ = TTOP$ + "\Otros Accesos Directos"
    End If
    '
      
    If OPMAX% < 1 Then GoTo 20
    '
    OPX% = ALTERNA%(TTOP$)
    If OPX% > OPMAX% Then
        GoTo 20
    ' '\\\OT-07-0386-FL-17/10/07///
      ElseIf OPX% < 1 Then
        GoTo 99
    End If
      '\\\OT-07-0386-FL-16/10/07///
    '
    If OPX% = OPTRAZA% Then
        If DERACCE%("TRAZORF", "Registro de Trazabilidad de O/F's") = 2 Then
          Call CONECRUN("TRAZAB07/OFABRIC", "Registro de Trazabilidad de O/F's")
        End If
        GoTo 99
      ElseIf OPX% = OPREPA% Then
        If DERACCE%("REPARA", "Sub-Sistema de Reparaciones") = 2 Then
          Call CONECRUN("REPARA07", "Sub-Sistema de Reparaciones")
        End If
        GoTo 99
        '\\\OT-07-0386-FL-17/10/07///
      ElseIf OPX% = OPSIMU% Then
        If DERACCE%("SIMULA07", "Simulación de Restricciones de Stock") = 2 Then
          Call CONECRUN("SIMULA07", "Simulación de Restricciones de Stock")
        End If
        GoTo 99
        '\\\OT-07-0386-FL-17/10/07///
          '
    End If
    '
20  Call ACCDIR("ORDEFABR")
    '
99  Command8.Enabled = True
End Sub


Private Sub Command45_Click()
    Call ORFANU
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    OPX% = COMALTER%("Opciones de Trabajo con O/F's:\\Cierre de O/F's\Re-Apertura\Anular Cierre de O/F")
    If OPX% = 1 Then
       If EXISTE%("TRAZAB07.EXE") > 0 Then
          If DERACCE%("CIERROFT", "Cierre de O/F's con Trazabilidad") = 2 Then
             Call CONECRUN("TRAZAB07/ORFCIE", "Cierre de O/F's con Trazabilidad")
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
    End If
99  Command46.Enabled = True
End Sub


Private Sub ConfGral_Click()
Call Command14_Click
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
    VERANTER$ = "ORDEFABR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If APLINVO$ = "OPRESERV" Then
      RESERV07.Show 1
      Call FINAL("")
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
'    If EXISTE%("REPARA02.EXE") < 1 Then
'      If EXISTE%("REPABO02.EXE") < 1 Then
'        Command3.Enabled = False
'        Label5.Enabled = False
'      End If
'    End If
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
   XX$ = REGSEL$("SECPROD/Sector Productivo Destino")
   If Len(XX$) <= 4 Then
      Exit Sub
   End If
   '
   SECCI = CVS(Left$(XX$, 4))
   SecProd$ = Mid$(XX$, 5, 32)
   DEPODEST% = Asc(Mid$(XX$, 81, 1))
   If DEPODEST% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%))) = "" Then
     Call MENSERR(24, "Falta Definir Depósito Asociado\al Sector '" + TRIM$(SecProd$) + "'.")
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
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + DEPOSI$ + " a " + SecProd$
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
         REFE$ = "TR.a " + SecProd$
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
    Call ABREORFAB
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
    MONUME$ = "AUTO"                   ' modo de numeracion
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("ORFAEMI")
    '
    X1$ = ""
100 GoSub ARMAPRIMERA
    X1$ = OBJPLA$("ORFANUE-1", X0$)
    If X1$ = "" Then GoTo 999
    '
    GoSub VERICODI
    If CIOF% < 1 Then GoTo 100
    CIXI% = CIOF%
    SPX$ = String$(4, 0): SPY$ = ""
    '
    GoSub AJUCANORF
    GoSub ARMASEGUNDA
    TMT% = TIMATE%(CIOF%)   ' PARA PONER TIPO DE MATERIAL EN POSICION 257,1
    XX22$ = X1$ + Chr$(TMT%)
102 XX22$ = OBJPLA$("ORFANUE-4", XX22$)
    REFOF$ = TRIM$(Mid$(XX22$, 258, 30))
    REF2$ = TRIM$(Mid$(XX22$, 116, 18))
    X2$ = Left$(XX22$, 256)
    If X2$ = "" Then
       If APLIXYZ$ <> "" Then
          Call FINAL("")
          Exit Sub
       End If
       GoTo 100
    End If
    '
    GoSub VALIDA1                   ' validacion de entrada
    If VALIDA% < 1 Then
       Call MENSERR(24, "Datos Incompletos o no Validos")
       GoTo 102
    End If
    '
103 XACCE$ = ""
    Call CIERRARCH("*.*")
    '
    TERENUME$ = ""
    If REPIM% <> 2 Then
       '
       NROPROX& = 1
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
       SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
       '
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
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("INDIORF")
    Call CIERRARCH("ECORDEP")
    Call CIERRARCH("ECORDEF")
    '
    Call COMUNI("Orden de Fabricación '" + NORFA$ + "' Generada.")
    '
    If APLIXYZ$ <> "" Then
       Call FINAL("")
       Exit Sub
    End If
    '
    GoTo 100
    '
' ************************************* subrutinas asociadas
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
       Call REPLA(X0$, MKI$(CODINT%(CODI6$)), 71, 2)
    End If
    '
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
    Call REPLA(YACCE$, Color$(CIOF%), 15, 16)
    '
    RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIOF%))
    If RFF$ <> "" Then
       REGEST& = CVS(Left$(RFF$, 4))
       CIHIL% = CVI(Mid$(REGLEIDO$("ESTRUNUE", REGEST&), 3, 2))
       Call REPLA(YACCE$, TIPOHIL$(CIHIL%), 63, 12)
       Call REPLA(YACCE$, Titulo$(CIHIL%), 75, 12)
       Call REPLA(YACCE$, Color$(CIHIL%), 87, 16)
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
    If LINDIVIS(CIOF%) >= 1 Then
       CANORF1 = (1 + Int((CANORF - 1) / LINDIVIS(CIOF%))) * LINDIVIS(CIOF%)
       If Abs(CANORF - CANORF1) > 0.5 Then
          CANORF = CANORF1
          Call MENSERR(24, "Cantidad Ajustada a Múltiplo\de Lote Indivisible.")
       End If
    End If
    '
    If LOSTA < CANORF Then
      XLMX$ = MKI$(CI%) + MKS$(CANORF)
      XLMY$ = OBJPLA$("CARLOREPOMA", XLMX$)
      If XLMY$ <> "" Then
        LOSTA = CVS(Mid$(XLMY$, 3, 4))
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
    CANOR = CVD(Mid$(X3$, 89, 8))
    If CANOR < 1 Then Return
    If CVS(Mid$(X3$, 97, 4)) < 1 Then Call REPLA(X3$, MKS$(CANOR), 97, 4)
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
    If REFOF$ <> "" Then DESCRI$ = TRIM$(REFOF$) + " - " + DESCRI$
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
    If REFOF$ <> "" Then
       Call REPLA(X$, REFOF$, 21, 20)
    End If
    If FCIE <> 0 Then Call REPLA(X$, FECHATEX$(FCIE), 43, 8)
    '
    TINDIORF$ = X$      ' prepara registro para INDIORF.DAT
    VALIDA% = 1
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
       FLEXCONN.Execute (FILTSQL$("DELETE  FROM ORDEFABR WHERE NumeOrfa = " & NRO & ""))
    End If
    '
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-1"
      NORFA$ = NBRS$
      II% = 1
      '
      SQLX$ = " SELECT * FROM OrdeFabr where NumeOrfa<1"
      Dim ORDEFFFABR As ADODB.Recordset
      Set ORDEFFFABR = New ADODB.Recordset
      ORDEFFFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockOptimistic, adCmdText
      'With OrdeFabr
      With ORDEFFFABR
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
          !NUMEORFA = NRO&
          !Referen = AJUSTI$(DESCRI$, 64)
          !IdSubOr = NORFA$
          !FechGen = Now
          !Cod_Inte = CIOF%
          !COD_EXTE = CODEXT$(CIOF%)
          !Descrip = DESCRIT0$(CIOF%)
          !Unidad = UNIMED$(CIOF%)
          !CanProy = CAN#
          If Len(DESTI$) > 8 Then
          DESTI$ = TRIM$(Left(DESTI$, 8))
          End If
          !Destino = DESTI$
          !FechEnSol = CVDAT(FENTRE%)
          !Priorid = PRIO%
          !Statuorf = Status%
          !FechLan = CVDAT(FLAN%)
          !SecciPro = SecProd
          !FechCierr = CVDAT(0)
          !CanApro = 0
          !CanRech = 0
          !IdenRetra = ""
          !FechProyEntre = CVDAT(0)
          !FechTempEntre = CVDAT(0)
          !FechRecal = CVDAT(0)
          !Observa = ""
          !ID_ORFA = REF2$
        .Update
      End With
      Set ORDEFFFABR = Nothing
      '
      CANS$ = " - (" + TRIM$(Str$(CAN#)) + ")"
      DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
      ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
      '
      Call FILTERPUTRECORD("ECORDEF", 1, NORFA$, ECORDEF$)
      If Status% = 1 Then
        Call FILTERPUTRECORD("ECORDEP", 1, NORFA$, ECORDEF$)
      End If
      '
Return
    '
999 Call CIERRARCH("*.*")
    Call FRESHALL
    Screen.MousePointer = 1
    If APLIXYZ$ <> "" Then
       Call FINAL("")
    End If
End Sub



Private Sub Frame9_DblClick()
  GENORF07.Show 1
End Sub

'
Private Sub GenReservOC_Click()
Call RESERV07.Command3_Click
End Sub

Private Sub Image2_DBLClick()
  '\\\OT-05-0733-WAR-14/11/05///
  'NUEORFA04.Show 1
  '\\\OT-05-0733-WAR-FIN///
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

Private Sub modiocomof_Click()
       VDEF$ = "": VDE1$ = ""
30     OB1$ = OBJPLA$("ORFAORCO", VDE1$)
       If OB1$ = "" Then Exit Sub
       '
       NORFA$ = Mid$(OB1$, 1, 12)
       CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
       OCOMP$ = VALOBADA("ORDEFABR", "ID_ORFA", CONDI$, "NOMESS")
       Call REPLA(VDEF$, NORFA$, 1, 12)
       Call REPLA(VDEF$, OCOMP$, 13, 18)
40     OBX$ = OBJPLA$("ORFAORC1", VDEF$)
       If OBX$ = "" Then
          VDE1$ = OB1$
          GoTo 30
       End If
       
       OCOMP$ = Mid$(OBX$, 13, 18)
       Call ACTUREGISTRO("ORDEFABR", "ID_ORFA", CONDI$, OCOMP$, REGAF&, "NOMESS")
       Call COMUNI("Referencia 2 Guardada")
       GoTo 30
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






Private Sub otraccdir_Click()
Call Command8_Click
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









Sub REPRINORFA()
   '
   Call ORFNUE(2)
   '
End Sub
'
'
Sub ANUCIEOF()
   '
   ' 1.- Selección del Producto
10 Call SELECHO("MASTER")
   If VALACT1$("MASTER") = "" Then Exit Sub
   CODD$ = VALACT1$("MASTER")
   CIXI% = CODINT%(CODD$)
   If CIXI% < 1 Then Exit Sub
   '
   ' 2.- Selección de la O/F
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND CodiMovi = 'OF' "
   SQLX$ = SQLX$ + "AND CantIngre > 0 "
   SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC "
   'Set Movisel = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim MOVISEL As ADODB.Recordset
   Set MOVISEL = New ADODB.Recordset
   MOVISEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
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
         Y$ = REGBLAN("ECORDEP")
         NOFA$ = !DoPriCor
         FERE$ = FECHATEX$(CVINT(!fechMov))
         CANX$ = FORMATNUM$(!cantingre, "F10.1")
         If Not IsNull(!FeReMovi) Then DIVE$ = FORMATNUM$(CDbl(!FeReMovi) - 0.00001, "F18.6")
         TTXX$ = FERE$ + "    " + CANX$ + "  " + DIVE$
         Call REPLA(Y$, NOFA$, 1, 12)
         Call REPLA(Y$, TTXX$, 13, 52)
         '
         JJ& = JJ& + 1
         Call GRAREG("!ECORDEP", Y$, JJ&)
         '
      .MoveNext
      Loop
   End With
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
   FREMOX$ = TRIM$(Mid$(RSEL$, 37, 18))
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE DoPriCor = '" & NOFASE$ & "' "
   SQLX$ = SQLX$ + "AND CDBL(FeReMovi) >= " & FREMOX$ & " "
   SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC, NuorItem ASC"
   'Set MOVISELE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim MOVISELE As ADODB.Recordset
    Set MOVISELE = New ADODB.Recordset
    MOVISELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

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
         If !NuOrItem < NUORAN% Then GoTo 39
         '.Edit
         REFA$ = !ReferCor
         CIXI% = !Cod_Inte
         NUORAN% = !NuOrItem
         !ReferCor = "Anul." + REFA$
         !cantingre = 0
         !CantSalid = 0
         .Update
         If CIXI% > 0 Then
            CIMOVI%(CIXI%) = 1
            IDLOTX$ = TRIM$(!IdenLote)
            If IDLOTX$ <> "" Then
               Call ACSALOTE(CIXI%, IDLOTX$)
            End If
         End If
      .MoveNext
      Loop
   End With
   '
39 Set MOVISELE = Nothing
   '
'3.- Actualiza saldo pendiente de la Orden de Fabricacion
   Call ABREORFAB
   SQLX$ = "SELECT * FROM OrdeFabr"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFASE$ & "'"
   Dim ORDEFFFAB As ADODB.Recordset
   Set ORDEFFFAB = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With ORDEFFFAB
     '.FindFirst "IdSubOr = '" & NOFASE$ & "'"
     'If .NoMatch = False Then
       '.Edit
     On Error Resume Next
     If Not (.EOF And .BOF) Then
      On Error GoTo 0
      !CanApro = !CanApro - CANUL
      !Statuorf = 1                ' MARCA COMO EN CURSO
      .Update
     End If
     On Error GoTo 0
   End With
   ORDEFFFAB.Close
   Set ORDEFFFAB = Nothing
   '
'   Call VESARRGLO
   Screen.MousePointer = 1
   '
   Call COMUNI("Movimiento Anulado")
   GoTo 10
   
End Sub
'
Sub REAPEROF()
    '
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
        Y$ = REBLA$
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", Y$, JJ&)
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
                !Statuorf = 1
                .Update
                Y$ = REBLA$
                Call REPLA(Y$, NOFA$, 1, 12)
                Call REPLA(Y$, ECOARCH$("ECORDEF", NOFA$), 13, 52)
                Call FILTERPUTRECORD("ECORDEP", 1, NOFA$, Y$)
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
