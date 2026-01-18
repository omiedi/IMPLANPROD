VERSION 5.00
Begin VB.Form INFOFA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Producción"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10515
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9450
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   61
      Top             =   2415
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command32 
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
      Left            =   9450
      Picture         =   "INFOFA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Control de Base de Datos de Operaciones Pendientes"
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Operar."
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
      Left            =   9450
      Picture         =   "INFOFA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Padron de Operarios"
      Top             =   3990
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
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
      Height          =   1800
      Left            =   210
      TabIndex        =   44
      Top             =   3780
      Width           =   4425
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   45
         Top             =   945
         Width           =   2745
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
            Picture         =   "INFOFA.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "n de Hojas de Ruta"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencia de Rutas  y Operaciones"
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
            Top             =   105
            Width           =   1905
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
         TabIndex        =   51
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command18 
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
            TabIndex        =   57
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
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
            TabIndex        =   56
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
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
            TabIndex        =   55
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
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
            TabIndex        =   54
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   53
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   52
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   48
         Top             =   420
         Width           =   2640
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
            Picture         =   "INFOFA.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Acceso Directo a Maestro de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
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
            Left            =   945
            TabIndex        =   50
            Top             =   75
            Width           =   1800
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
      Left            =   9450
      Picture         =   "INFOFA.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1365
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
      Left            =   9450
      Picture         =   "INFOFA.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cerrar - Salir de la Aplicación"
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
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
      Height          =   1800
      Left            =   210
      TabIndex        =   2
      Top             =   1890
      Width           =   4425
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   28
         Top             =   945
         Width           =   2745
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
            Picture         =   "INFOFA.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Acceso Directo a Programación CRP / MRP-II"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Programacion de Carga de Máquina"
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
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   25
         Top             =   420
         Width           =   2640
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
            Picture         =   "INFOFA.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Acceso Directo a Ordenes de Fabricación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "B-Datos de Ordenes de Fabricación"
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
            TabIndex        =   27
            Top             =   80
            Width           =   1905
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
      BackColor       =   &H00C0FFC0&
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
      Height          =   5265
      Left            =   4830
      TabIndex        =   1
      Top             =   315
      Width           =   4425
      Begin VB.CommandButton Command33 
         Caption         =   "Command33"
         Height          =   615
         Left            =   2640
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   1905
         Left            =   3465
         ScaleHeight     =   1845
         ScaleWidth      =   690
         TabIndex        =   64
         Top             =   420
         Width           =   750
         Begin VB.CommandButton Command31 
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
            Left            =   0
            Picture         =   "INFOFA.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Acceso a Productividad General e Individual por Operario"
            Top             =   0
            Width           =   690
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Control de Mano de Obra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   30
            TabIndex        =   65
            Top             =   800
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Productos en Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1220
         Left            =   210
         TabIndex        =   39
         Top             =   3885
         Width           =   4005
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   40
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
               Picture         =   "INFOFA.frx":199A
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Listado por Código de Producto"
               Top             =   0
               Width           =   590
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
               Picture         =   "INFOFA.frx":1CA4
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Listado por Orden de Fabricación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de Avance Prod.en Proceso"
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
               TabIndex        =   43
               Top             =   105
               Width           =   1695
            End
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   13
         Top             =   1680
         Width           =   2010
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
            Left            =   -20
            Picture         =   "INFOFA.frx":1FAE
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Consulta por Código de Producto"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Código de Producto"
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
            TabIndex        =   15
            Top             =   105
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   16
         Top             =   1260
         Width           =   2010
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
            Picture         =   "INFOFA.frx":22B8
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Consulta por Número de O/F"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Nro. O/F"
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
            TabIndex        =   18
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   19
         Top             =   840
         Width           =   2010
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
            Left            =   0
            Picture         =   "INFOFA.frx":25C2
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Informe por Máquina / Equipo y Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Máquina"
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
            TabIndex        =   21
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   22
         Top             =   420
         Width           =   2010
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
            Picture         =   "INFOFA.frx":28CC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Consulta de Produccion por Fecha"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Fecha"
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
            TabIndex        =   24
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1220
         Left            =   210
         TabIndex        =   12
         Top             =   2520
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   35
            Top             =   420
            Width           =   3165
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
               Picture         =   "INFOFA.frx":2A16
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Consulta Pendientes por Orden de fabricación"
               Top             =   0
               Width           =   560
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
               Left            =   2520
               Picture         =   "INFOFA.frx":2D20
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Consulta de Pendientes por Máquina / Equipo"
               Top             =   0
               Width           =   590
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Informes Parciales de Producción"
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
               TabIndex        =   37
               Top             =   105
               Width           =   1695
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Partes de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   4425
      Begin VB.ListBox SORTLIST 
         Height          =   255
         Left            =   1155
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   1155
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   3435
         TabIndex        =   31
         Top             =   525
         Width           =   3500
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
            Left            =   2310
            Picture         =   "INFOFA.frx":302A
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Partes de Fabricación Operaciones Pendientes por O/F"
            Top             =   0
            Width           =   590
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
            Picture         =   "INFOFA.frx":3334
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Partes de Fabricación por Programacion Diaria"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command30 
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
            Left            =   2900
            Picture         =   "INFOFA.frx":363E
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Informe de Tiempos de Operaciones Cumplidas"
            Top             =   0
            Width           =   590
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes Parciales de Producción"
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
            Width           =   1695
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8900
      Picture         =   "INFOFA.frx":3788
      Top             =   5040
      Width           =   2010
   End
End
Attribute VB_Name = "INFOFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
10  Call SELECHO("PADMAQ")
    VDEVU$ = VALACT1$("PADMAQ")
    If VDEVU$ <> "" Then
      MAQUISEL$ = AJUSTI$(VDEVU$, 6)
      Call DESHASFECHA(DES%, HAS%, PERIO$)
      If PERIO$ <> "" Then
        SORTLIST.Clear
        Call BLOQARCH("INFOPEFA")
        JJ& = 0
        '
        Screen.MousePointer = 11
        FIN& = ULTREG&("OPERFAB")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          ZZ$ = REGLEIDO$("OPERFAB", U&)
          If CVI(Mid$(ZZ$, 157, 2)) >= DES% Then
            If CVI(Mid$(ZZ$, 157, 2)) <= HAS% Then
              MAQUI$ = Mid$(ZZ$, 151, 6)
              If MAQUI$ = MAQUISEL$ Then
                FESOR$ = CSTRING$(Mid$(ZZ$, 157, 2), 1, 6, 0, 2)
                HORTEX$ = AJUSTD$(HORATEX$(CVI(Mid$(ZZ$, 159, 2))), 6)
                RESOR$ = CSTRING$(MKS$(U&), 3, 6, 0, 2)
                SORTEX$ = MAQUI$ + FESOR$ + HORTEX$ + RESOR$
                SORTLIST.AddItem SORTEX$
              End If
            End If
          End If
        Next U&
        '
        For KK% = 1 To SORTLIST.ListCount
          REOPE& = Val(Mid$(SORTLIST.List(KK% - 1), 19, 6))
          ZZ$ = REGLEIDO$("OPERFAB", REOPE&)
          XX$ = REGBLAN$("INFOPEFA")
          NORFA$ = Left$(ZZ$, 12)
          NOPE% = CVI(Mid$(ZZ$, 15, 2))
          DOPE$ = Mid$(ZZ$, 17, 48)
          MAQUI$ = Mid$(ZZ$, 151, 6)
          FEOP% = CVI(Mid$(ZZ$, 157, 2))
          HOOP% = CVI(Mid$(ZZ$, 159, 2))
          CIIX% = CVI(Mid$(ZZ$, 13, 2))
          CAREA# = CVD(Mid$(ZZ$, 87, 8))
          TEMPLE = 60 * CVS(Mid$(ZZ$, 161, 4)) ' PASA A SEGUNDOS
          NUOROP% = Asc(Mid$(ZZ$, 67, 1))
          TSTAN = 60 * TIESTAN(CIIX%, NOPE%, CAREA#, NUOROP%) ' PASA A SEGUNDOS
          RENDIM = 999: If TEMPLE > 0.01 Then RENDIM = 100 * TSTAN / TEMPLE
          If RENDIM > 999 Then RENDIM = 999
          '
          Call REPLA(XX$, NORFA$, 3, 12)
          Call REPLA(XX$, MKI$(NOPE%), 15, 2)
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, MAQUI$, 65, 6)
          Call REPLA(XX$, MKI$(FEOP%), 71, 2)
          Call REPLA(XX$, MKI$(HOOP%), 73, 2)
          Call REPLA(XX$, NORFA$, 75, 12)
          Call REPLA(XX$, MKI$(CIIX%), 87, 2)
          Call REPLA(XX$, MKI$(NOPE%), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(CAREA#), 139, 8)
          Call REPLA(XX$, MKS$(TEMPLE), 147, 4)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          '
          For KL% = 165 To 172 Step 2
            LEGA% = CVI(Mid$(ZZ$, KL%, 2))
            If LEGA% > 0 Then
              If LEGA% <> 8224 Then
                Call REPLA(XX$, MKI$(LEGA%), 1, 2)
                Call REPLA(XX$, MKI$(LEGA%), 167, 2)
                JJ& = JJ& + 1
                Call GRAREG("INFOPEFA", XX$, JJ&)
              End If
            End If
          Next KL%
          '
        Next KK%
        '
        Call SETULTREG("INFOPEFA", JJ&)
        Call CIERRARCH("INFOPEFA")
        '
        Call HEADERS("INFOPEFA", "")
        Call HEADERS("INFOPEFA", "Corresponde a: " + TRIM$(MAQUISEL$) + " - Desde " + FECHATEX$(DES%) + " hasta " + FECHATEX$(HAS%))
        Call CONECRUN("*INFAMAQ", "Informes de Producción por Equipo")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    'Call FINAL("*COINFEQ")
    Command15.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*PADRON")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GESALPROP(1)
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
    ACONEC$ = "CARMAQ02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
      ACONEC$ = "CARGMAQ"
    End If
    Call CONECRUN(ACONEC$, "")
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
    Command26.Enabled = False
10  Call SELECHO("FECHA")
    VDEVU$ = VALACT1$("FECHA")
    If VDEVU$ <> "" Then
      FF% = FECHANUM(VDEVU$)
      SORTLIST.Clear
      Call BLOQARCH("INFOPEFA")
      JJ& = 0
      If FF% > 0 Then
        Screen.MousePointer = 11
        FIN& = ULTREG&("OPERFAB")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          ZZ$ = REGLEIDO$("OPERFAB", U&)
          If CVI(Mid$(ZZ$, 157, 2)) = FF% Then
            MAQUI$ = Mid$(ZZ$, 151, 6)
            FESOR$ = CSTRING$(Mid$(ZZ$, 157, 2), 1, 6, 0, 2)
            HORTEX$ = AJUSTD$(HORATEX$(CVI(Mid$(ZZ$, 159, 2))), 6)
            RESOR$ = CSTRING$(MKS$(U&), 3, 6, 0, 2)
            SORTEX$ = MAQUI$ + FESOR$ + HORTEX$ + RESOR$
            SORTLIST.AddItem SORTEX$
          End If
        Next U&
        '
        For KK% = 1 To SORTLIST.ListCount
          REOPE& = Val(Mid$(SORTLIST.List(KK% - 1), 19, 6))
          ZZ$ = REGLEIDO$("OPERFAB", REOPE&)
          XX$ = REGBLAN$("INFOPEFA")
          NORFA$ = Left$(ZZ$, 12)
          NOPE% = CVI(Mid$(ZZ$, 15, 2))
          DOPE$ = Mid$(ZZ$, 17, 48)
          MAQUI$ = Mid$(ZZ$, 151, 6)
          FEOP% = CVI(Mid$(ZZ$, 157, 2))
          HOOP% = CVI(Mid$(ZZ$, 159, 2))
          CIIX% = CVI(Mid$(ZZ$, 13, 2))
          CAREA# = CVD(Mid$(ZZ$, 87, 8))
          TEMPLE = 60 * CVS(Mid$(ZZ$, 161, 4)) ' PASA A SEGUNDOS
          NUOROP% = Asc(Mid$(ZZ$, 67, 1))
          TSTAN = 60 * TIESTAN(CIIX%, NOPE%, CAREA#, NUOROP%) ' PASA A SEGUNDOS
          RENDIM = 999: If TEMPLE > 0.01 Then RENDIM = 100 * TSTAN / TEMPLE
          If RENDIM > 999 Then RENDIM = 999
          '
          Call REPLA(XX$, NORFA$, 3, 12)
          Call REPLA(XX$, MKI$(NOPE%), 15, 2)
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, MAQUI$, 65, 6)
          Call REPLA(XX$, MKI$(FEOP%), 71, 2)
          Call REPLA(XX$, MKI$(HOOP%), 73, 2)
          Call REPLA(XX$, NORFA$, 75, 12)
          Call REPLA(XX$, MKI$(CIIX%), 87, 2)
          Call REPLA(XX$, MKI$(NOPE%), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(CAREA#), 139, 8)
          Call REPLA(XX$, MKS$(TEMPLE), 147, 4)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          '
          For KL% = 165 To 172 Step 2
            LEGA% = CVI(Mid$(ZZ$, KL%, 2))
            If LEGA% > 0 Then
              If LEGA% <> 8224 Then
                Call REPLA(XX$, MKI$(LEGA%), 1, 2)
                Call REPLA(XX$, MKI$(LEGA%), 167, 2)
                JJ& = JJ& + 1
                Call GRAREG("INFOPEFA", XX$, JJ&)
              End If
            End If
          Next KL%
          '
        Next KK%
        '
        Call SETULTREG("INFOPEFA", JJ&)
        Call CIERRARCH("INFOPEFA")
        '
        FEX = FF%
        Call HEADERS("INFOPEFA", "")
        Call HEADERS("INFOPEFA", "Corresponde a: " + FECHATEX$(FEX))
        Call CONECRUN("*INFADIA", "Informes de Producción por Fecha")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    'Call FINAL("*COINFDI")
    Command26.Enabled = True
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    Call FINAL("*COINFOF")
    Command27.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
10  Screen.MousePointer = 1
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
      CIIX% = CODINT%(VDEVU$)
      SORTLIST.Clear
      Call BLOQARCH("INFOPEFA")
      JJ& = 0
      If CIIX% > 0 Then
        Screen.MousePointer = 11
        RFF$ = RECFILT$("OPERFAB", 2, MKI$(CIIX%))
        FIN& = ULTREG&("OPERFAB")
        For KH& = 1 To Len(RFF$) Step 4
          U& = CVS(Mid$(RFF$, KH&, 4))
          If U& > 0 Then
            If U& <= FIN& Then
              ZZ$ = REGLEIDO$("OPERFAB", U&)
              'If Mid$(ZZ$, 177, 6) <> String$(6, 0) Then
                If CVI(Mid$(ZZ$, 13, 2)) = CIIX% Then
                  NORFA$ = Mid$(ZZ$, 1, 12)
                  NOPER$ = CSTRING$(Mid$(ZZ$, 15, 2), 1, 6, 0, 2)
                  FESOR$ = CSTRING$(Mid$(ZZ$, 157, 2), 1, 6, 0, 2)
                  HORTEX$ = AJUSTD$(HORATEX$(CVI(Mid$(ZZ$, 159, 2))), 6)
                  RESOR$ = CSTRING$(MKS$(U&), 3, 6, 0, 2)
                  SORTEX$ = NORFA$ + NOPER$ + FESOR$ + HORTEX$ + RESOR$
                  SORTLIST.AddItem SORTEX$
                End If
              'End If
            End If
          End If
        Next KH&
        '
        If SORTLIST.ListCount < 1 Then
          Call MENSERR(24, "Código " + TRIM$(CODEXT$(CIIX%)) + " sin Operaciones Registradas.")
          GoTo 10
        End If
        '
        For KK% = 1 To SORTLIST.ListCount
          REOPE& = Val(Mid$(SORTLIST.List(KK% - 1), 31, 6))
          ZZ$ = REGLEIDO$("OPERFAB", REOPE&)
          XX$ = REGBLAN$("INFOPEFA")
          NORFA$ = Left$(ZZ$, 12)
          NOPE% = CVI(Mid$(ZZ$, 15, 2))
          DOPE$ = Mid$(ZZ$, 17, 48)
          MAQUI$ = Mid$(ZZ$, 151, 6)
          FEOP% = CVI(Mid$(ZZ$, 157, 2))
          HOOP% = CVI(Mid$(ZZ$, 159, 2))
          CIIX% = CVI(Mid$(ZZ$, 13, 2))
          CAREA# = CVD(Mid$(ZZ$, 87, 8))
          TEMPLE = 60 * CVS(Mid$(ZZ$, 161, 4)) ' PASA A SEGUNDOS
          NUOROP% = Asc(Mid$(ZZ$, 67, 1))
          TSTAN = 60 * TIESTAN(CIIX%, NOPE%, CAREA#, NUOROP%) ' PASA A SEGUNDOS
          RENDIM = 999: If TEMPLE > 0.01 Then RENDIM = 100 * TSTAN / TEMPLE
          If RENDIM > 999 Then RENDIM = 999
          '
          Call REPLA(XX$, NORFA$, 3, 12)
          Call REPLA(XX$, MKI$(NOPE%), 15, 2)
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, MAQUI$, 65, 6)
          Call REPLA(XX$, MKI$(FEOP%), 71, 2)
          Call REPLA(XX$, MKI$(HOOP%), 73, 2)
          Call REPLA(XX$, NORFA$, 75, 12)
          Call REPLA(XX$, MKI$(CIIX%), 87, 2)
          Call REPLA(XX$, MKI$(NOPE%), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(CAREA#), 139, 8)
          Call REPLA(XX$, MKS$(TEMPLE), 147, 4)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          '
          For KL% = 165 To 172 Step 2
            LEGA% = CVI(Mid$(ZZ$, KL%, 2))
            If LEGA% > 0 Then
              If LEGA% <> 8224 Then
                Call REPLA(XX$, MKI$(LEGA%), 1, 2)
                Call REPLA(XX$, MKI$(LEGA%), 167, 2)
                JJ& = JJ& + 1
                Call GRAREG("INFOPEFA", XX$, JJ&)
              End If
            End If
          Next KL%
          '
        Next KK%
        '
        Call SETULTREG("INFOPEFA", JJ&)
        Call CIERRARCH("INFOPEFA")
        '
        Call HEADERS("INFOPEFA", "")
        Call HEADERS("INFOPEFA", "Codigo: " + TRIM$(CODEXT$(CIIX%)) + " - " + TRIM$(DESCRIT0$(CIIX%)))
        Call CONECRUN("*INFACOD", "Informes de Producción por Articulo")
        Screen.MousePointer = 1
        GoTo 10
        '
      End If
    End If
    Command28.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("INFOFAB")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call GESALPROP(2)
   Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   Command30.Enabled = False
   If DERACCE%("MODINFOP", "Modificar Partes de Fabricacion") > 0 Then
      Call INFOPER
   End If
   Command30.Enabled = True
End Sub

Private Sub Command31_Click()
    '
    Command31.Enabled = False
    If EXISTE%("EFIMOB02.EXE") > 0 Then
         Call CONECRUN("EFIMOB02", "Control de Productividad de Mano de Obra")
       Else
         Call CONECRUN("EFIMOBRA", "Control de Productividad de Mano de Obra")
    End If
    Command31.Enabled = True
Exit Sub
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MMXX$ = VALACT1$("SELMESA")
    CORRESP$ = MESANO$(MMXX$, DES%, HAS%)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OPERFAB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("OPERFAB", U&)
       FFIN% = CVI(Mid$(X$, 157, 2))
       If FFIN% >= DES% Then
         If FFIN% <= HAS% Then
           LEGOPER$ = Mid$(X$, 165, 8)
           HFIN% = CVI(Mid$(X$, 159, 2))
           NOFA$ = Mid$(X$, 1, 12)
           CIXI% = CVI(Mid$(X$, 13, 2))
           NOPE% = CVI(Mid$(X$, 15, 2))
           DOPE$ = Mid$(X$, 17, 48)
           CAREA# = CVD(Mid$(X$, 87, 8))
           TREAL = 60 * CVS(Mid$(X$, 161, 4)) ' PASA A SEGUNDOS
           TSTAN = 60 * CVS(Mid$(X$, 173, 4)) ' PASA A SEGUNDOS
           COEFI = 999: If TSTAN > 0.1 Then COEFI = 100 * TREAL / TSTAN
           If COEFI > 999 Then COEFI = 999
           '
           ZZ1$ = REGBLAN$("COMOBRA")
           Call REPLA(ZZ1$, NOFA$, 3, 12)
           Call REPLA(ZZ1$, MKI$(CIXI%), 15, 2)
           Call REPLA(ZZ1$, MKI$(NOPE%), 17, 2)
           Call REPLA(ZZ1$, DOPE$, 19, 48)
           Call REPLA(ZZ1$, MKD$(CAREA#), 67, 8)
           Call REPLA(ZZ1$, MKI$(FFIN%), 75, 2)
           Call REPLA(ZZ1$, MKI$(HFIN%), 77, 2)
           Call REPLA(ZZ1$, MKS$(TREAL), 81, 4)
           Call REPLA(ZZ1$, MKS$(TSTAN), 85, 4)
           Call REPLA(ZZ1$, MKS$(COEFI), 89, 4)
           '
           For KK% = 1 To Len(LEGOPER$) Step 2
             LEGA% = CVI(Mid$(LEGOPER$, KK%, 2))
             If LEGA% > 0 Then
               If LEGA% <> 8224 Then
                 ZZ$ = ZZ1$
                 Call REPLA(ZZ$, MKI$(LEGA%), 1, 2)
                 Call REPLA(ZZ$, MKI$(LEGA%), 79, 2)
                 JJ& = JJ& + 1
                 Call GRAREG("COMOBRA", ZZ$, JJ&)
               End If
             End If
           Next KK%
         End If
       End If
    Next U&
    '
    Call SETULTREG("COMOBRA", JJ&)
    Call CIERRARCH("COMOBRA")
    '
    Call FINAL("*SOCOMOB")
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command32_Click()
    Command32.Enabled = False
    USXX% = USNBR% Mod 100
    If USXX% <> 1 Then Exit Sub
    Screen.MousePointer = 11
    '
    Dim MAQUI$(512), MAEQUI$(512)
10  NUMAQ% = ULTREG&("PADMAQ")
    If NUMAQ% > 512 Then
        Call COMUNI("Demasiadas Máquinas Definidas.\Consulte al Soporte de Sistemas.")
        Exit Sub
    End If
    '
    For U& = 1 To NUMAQ%
       REGMA$ = REGLEIDO$("PADMAQ", U&)
       MAQUI$(U&) = Left$(REGMA$, 6)
       VECTOMAQ$ = VECTOMAQ$ + "@@" + Left$(REGMA$, 6)
       MAEQUI$(U&) = Mid$(REGMA$, 43, 6)
    Next U&
    '
    ACONTROL$ = "OPERFAB"
12  FIN& = ULTREG&(ACONTROL$)
    ATX% = 0
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$(ACONTROL$, U&)
       CIXI% = CVI(Mid$(XX$, 13, 2))
       NOPE% = CVI(Mid$(XX$, 15, 2))
       MAQUIASI$ = Mid$(XX$, 73, 6)
       MAQUISTA$ = MAQUISTAN$(CIXI%, NOPE%)
       If MAQUISTA$ = "" Then GoTo 19
       VUELTA% = 0
15     If MAQUISTA$ = MAQUIASI$ Then GoTo 19
       PQX% = InStr(VECTOMAQ$, "@@" + MAQUISTA$)
       If PQX% < 1 Then GoTo 19
       IMAQUI& = (PQX% + 7) / 8
       If IMAQUI& > NUMAQ% Then GoTo 19
       If TRIM$(MAEQUI$(IMAQUI&)) <> "" Then
          MAQUISTA$ = MAEQUI$(IMAQUI&)
          VUELTA% = VUELTA% + 1
          If VUELTA% < 6 Then GoTo 15
       End If
       Call REPLA(XX$, MAQUISTAN$(CIXI%, NOPE%), 73, 6)
       Call GRAREG(ACONTROL$, XX$, U&)
       ATX% = 1
19  Next U&
    Call CIERRARCH("*.*")
    If ATX% > 0 Then
         Call COMUNI("Inconsistencias Recuperadas.\Realice Cálculo de Carga de Máquinas.")
       Else
         Call COMUNI("No se Detectaron Inconsistencias Recuperables.")
    End If
    Screen.MousePointer = 1
    Command32.Enabled = True
End Sub

Private Sub Command33_Click()
'
FIN& = ULTREG&("OPERFAB")
Call DESHASFECHA(DES%, HAS%, PERIO$)
For U& = 1 To FIN&
  Call TRACE(20, U&, FIN&)
  X$ = REGLEIDO$("OPERFAB", U&)
  FF% = CVI(Mid$(X$, 157, 2))
  If FF% >= DES% And FF% <= HAS% Then
    JJ& = JJ& + 1
    Call GRAREG("OPERFA1", X$, JJ&)
  End If
Next U&
'
Call SETULTREG("OPERFA1", JJ&)
Call CIERRARCH("*.*")
'
Call FINAL("*OPERFA1")
'
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call FINAL("*SECAROF")
    Command4.Enabled = True
End Sub
'

Private Sub Command5_Click()
    Command4.Enabled = False
    Screen.MousePointer = 11
    On Error Resume Next
    Kill LUDAT$ + "OPERFAH.DAT"
    On Error GoTo 0
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
    Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
Call MENSERR(24, "No Disponible por el Momento")
Call INFAGEN
    'Call INFADIA
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call INFAGEN
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If EXISTE%("AMASTO02.EXE") > 0 Then
        Call CONECRUN("AMASTO02", "Maestro de Artículos de Stock")
      Else
        Call CONECRUN("HORUTAS", "Hojas de Ruta - Secuencia de Operaciones")
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("AMASTO02", "Maestro de Artículos de Stock")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    VERANTER$ = "INFOFAB"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("EFIMOB02.EXE") < 1 Then
      If EXISTE%("EFIMOBRA.EXE") < 1 Then
        Command31.Enabled = False
        Label11.Enabled = False
      End If
    End If
    '
End Sub

Sub INFAGEN()      ' INFORME DE PRODUCCION - GENERAL
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "INFORMES DE PRODUCCION - GENERALES"

100  NROOF$ = ""

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
     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
     RGFSX$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
     RGFSY$ = ""
     For KJ& = 1 To Len(RGFSX$) Step 4
       REOPE& = CVS(Mid$(RGFSX$, KJ&, 4))
       If REOPE& > 0 Then
         If REOPE& <= ULTREG&("OPERFAB") Then
           XX$ = REGLEIDO$("OPERFAB", REOPE&)
           If Left$(XX$, 12) = NROOF$ Then
             If Mid$(XX$, 177, 6) = String$(6, 0) Then ' aun pendiente
               RGFSY$ = RGFSY$ + MKS$(REOPE&)
             End If
           End If
         End If
       End If
     Next KJ&
     '
     If Len(RGFSY$) < 4 Then
       For KJ& = 1 To ULTREG&("OPERFAB")
         XX$ = REGLEIDO$("OPERFAB", KJ&)
         If Left$(XX$, 12) = NROOF$ Then
           If Mid$(XX$, 177, 6) = String$(6, 0) Then ' aun pendiente
             Call REPLA(XX$, MKI$(BVOF%), 65, 2)
             Call GRAREG("OPERFAB", XX$, KJ&)
             RGFSY$ = RGFSY$ + MKS$(KJ&)
           End If
         End If
       Next KJ&
       Call CIERRARCH("OPERFAB")
     End If
     '
     If Len(RGFSY$) < 4 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     J% = 0
     Call FRESHECHO("!OPERINF")
     SORTLIST.Clear
     For K% = 1 To Len(RGFSY$) Step 4
        REGI& = CVS(Mid$(RGFSY$, K%, 4))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        NOPE% = CVI(Mid$(XX$, 15, 2))
        TSOR$ = CSTRING$(Mid$(XX$, 15, 2), 1, 6, 8, 2) + Str$(REGI&)
        SORTLIST.AddItem TSOR$
     Next K%
     '
     For K% = 1 To SORTLIST.ListCount
        REGI& = Val(Mid$(SORTLIST.List(K% - 1), 7))
        XX$ = REGLEIDO$("OPERFAB", REGI&)
        CAPROD# = CVD(Mid$(XX$, 79, 8))
        CAREA# = CVD(Mid$(XX$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
        '
        If SALDO# > 0.05 Then
            '
            CI% = CVI(Mid$(XX$, 13, 2))
            NOPE% = CVI(Mid$(XX$, 15, 2))
            DOPE$ = Mid$(XX$, 17, 48)
            CAINF# = SALDO#
            MAQUI$ = Mid$(XX$, 73, 6)
            FEOP% = CVI(Mid$(XX$, 139, 2))
            HOOP% = CVI(Mid$(XX$, 141, 2))
            MALIB$ = "S"
            TITOT = CVS(Mid$(XX$, 131, 4))
            '
            J% = J% + 1
            YY$ = MKI$(J%) + Space$(58) + MKS$(REGI&)
            NOPX$ = AJUSTD$(Str$(NOPE%), 5) + " - " + AJUSTI$(TRIM$(Left$(DOPE$, 31)) + " " + MAQUI$, 38)
            NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            JJ& = J%
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
        End If
     Next K%
     '
     If J% < 1 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Orden de Fabricación.")
         Screen.MousePointer = 1
         GoTo 1405
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para O/F " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
         GoTo 1405
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     REGI& = CVS(Mid$(XX$, 61, 4))
     Call CARGAINF(REGI&)
     If REGI& < 1 Then
         GoTo 1415
       Else
         GoTo 1405
     End If
     '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
