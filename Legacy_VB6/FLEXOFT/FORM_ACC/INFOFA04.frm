VERSION 5.00
Begin VB.Form INFOFA04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Producción"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9945
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00006000&
      Height          =   6315
      Left            =   9030
      ScaleHeight     =   6255
      ScaleWidth      =   1110
      TabIndex        =   52
      Top             =   0
      Width           =   1170
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Acceso Directo a Maestro de Items de Stock"
         Top             =   1995
         Width           =   650
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
         Picture         =   "INFOFA04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
         Top             =   420
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
         Height          =   630
         Left            =   105
         Picture         =   "INFOFA04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1060
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Padron de Operarios"
         Top             =   2640
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Control de Integridad de la Base de Datos de Operaciones Pendientes"
         Top             =   3570
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   53
         Top             =   4410
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "INFOFA04.frx":0D72
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4725
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Mano de Obra"
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
      Left            =   210
      TabIndex        =   42
      Top             =   3780
      Width           =   4110
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   60
         Top             =   525
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "INFOFA04.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Acceso a Productividad General e Individual por Operario"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Eficiencia de Mano de Obra"
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
            TabIndex        =   62
            Top             =   105
            Width           =   1500
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
         TabIndex        =   43
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
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
      Height          =   1800
      Left            =   210
      TabIndex        =   2
      Top             =   1890
      Width           =   4110
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   26
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
            Picture         =   "INFOFA04.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   23
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
            Picture         =   "INFOFA04.frx":32A8
            Style           =   1  'Graphical
            TabIndex        =   24
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
            TabIndex        =   25
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   4410
      TabIndex        =   1
      Top             =   315
      Width           =   4425
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Productos en Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1220
         Left            =   210
         TabIndex        =   37
         Top             =   3465
         Width           =   4005
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   350
            ScaleHeight     =   585
            ScaleWidth      =   3270
            TabIndex        =   38
            Top             =   420
            Width           =   3330
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
               Left            =   2690
               Picture         =   "INFOFA04.frx":35B2
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Ajuste de Saldos Productos en Proceso"
               Top             =   0
               Width           =   590
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
               Left            =   2100
               Picture         =   "INFOFA04.frx":36FC
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Control de Operaciones y Productos en Proceso"
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
               Picture         =   "INFOFA04.frx":3A06
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Control de Operaciones Pendientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Control  Produc- tos en Proceso"
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
               Width           =   1695
            End
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1785
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   11
         Top             =   1440
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
            Picture         =   "INFOFA04.frx":3D10
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listado  por Código de Producto y Rango de Fechas"
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
            TabIndex        =   13
            Top             =   105
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   14
         Top             =   1100
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
            Picture         =   "INFOFA04.frx":401A
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   16
            Top             =   60
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   17
         Top             =   760
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
            Picture         =   "INFOFA04.frx":4324
            Style           =   1  'Graphical
            TabIndex        =   18
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
            TabIndex        =   19
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   20
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
            Picture         =   "INFOFA04.frx":462E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Informe de Produccion por Fecha"
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
            TabIndex        =   22
            Top             =   60
            Width           =   1905
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Operaciones Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1220
         Left            =   210
         TabIndex        =   10
         Top             =   2205
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3105
            TabIndex        =   33
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
               Picture         =   "INFOFA04.frx":4778
               Style           =   1  'Graphical
               TabIndex        =   36
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
               Picture         =   "INFOFA04.frx":4A82
               Style           =   1  'Graphical
               TabIndex        =   34
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
               TabIndex        =   35
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
         Size            =   9.75
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
      Width           =   4110
      Begin VB.ListBox SORTLIST 
         Height          =   255
         Left            =   1155
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   1155
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   3015
         TabIndex        =   29
         Top             =   525
         Width           =   3075
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
            Left            =   2440
            Picture         =   "INFOFA04.frx":4D8C
            Style           =   1  'Graphical
            TabIndex        =   32
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
            Picture         =   "INFOFA04.frx":5096
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Informes de Proyecto"
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
            Left            =   3460
            Picture         =   "INFOFA04.frx":53A0
            Style           =   1  'Graphical
            TabIndex        =   50
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
            TabIndex        =   31
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11025
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu PaFab 
      Caption         =   "Partes de Fabricación"
      Begin VB.Menu ProgDia 
         Caption         =   "Por Informe de Proyecto"
      End
      Begin VB.Menu OpPenOF 
         Caption         =   "Por Orden de Fabricación"
      End
      Begin VB.Menu TiOpCum 
         Caption         =   "Re-Informe Operaciones Cumplidas"
      End
      Begin VB.Menu php0a 
         Caption         =   "-"
      End
      Begin VB.Menu CtrlOpPen 
         Caption         =   "Control de Integridad de la B-D"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu CoProFec 
         Caption         =   "Consulta de Producción por Fecha"
      End
      Begin VB.Menu InfMayRa 
         Caption         =   "Informe por Máquina y Rango de Fechas"
      End
      Begin VB.Menu CoNumOF 
         Caption         =   "Consulta por Número de O/F"
      End
      Begin VB.Menu CoCodPro 
         Caption         =   "Consulta por Código de Producto"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu OpePen 
         Caption         =   "Operaciones Pendientes"
         Begin VB.Menu OrdFabri 
            Caption         =   "Por Orden de Fabricación"
         End
         Begin VB.Menu MaEq 
            Caption         =   "Por Máquina / Equipo"
         End
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu PenProc 
         Caption         =   "Pendientes en Proceso"
         Begin VB.Menu OrdFabric 
            Caption         =   "Por Orden de Fabricación"
         End
         Begin VB.Menu CodProduc 
            Caption         =   "Por Código de Producto"
         End
      End
   End
   Begin VB.Menu AcceDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu BdDOrFa 
         Caption         =   "Ordenes de Fabricación"
      End
      Begin VB.Menu Progmaqui 
         Caption         =   "Programación Carga de Máquinas"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu MaArtStck 
         Caption         =   "Maestro de Articulos de Stock"
      End
      Begin VB.Menu SecRuOp 
         Caption         =   "Secuencia de Operaciones"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu PadOper 
         Caption         =   "Padron Maestro de Operarios"
      End
      Begin VB.Menu php7 
         Caption         =   "-"
      End
      Begin VB.Menu OtAccDirec 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu ayd 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INFOFA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AUTOCONTROL()
    '
    '27/03/07 (OT-07-0087)
    '
    FORIPRE$ = TRIM$(CONTROL$("IFABRICA", "FAUTCONT"))
    '
    Call DESHASFECHA(DES1%, HAS1%, PERIO$)
    VDEVU$ = PERIO$
    If VDEVU$ <> "" Then
       FEC1 = CDbl(CVDAT(DES1%))
       FEC2 = CDbl(CVDAT(HAS1%))
    End If
    '
    Call SELECHO("SECPROD")
    VDEVU1$ = VALACT1$("SECPROD")
    VDEUV2$ = VALACT2$("SECPROD")
    '
    Call ABREORFAB
    sqlx$ = " SELECT * FROM InfoFab "
    sqlx$ = sqlx$ + " WHERE  INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
    sqlx$ = sqlx$ + " AND  INT(CDBL(FeHorLib))  <=  " & FEC2 & " "
    sqlx$ = sqlx$ + " ORDER BY INT(FeHorLib) ASC, Hora_Des ASC "
    On Error Resume Next
    Set PROXFECTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Registros no Procesables.\Debe Ejecutar Opción de Control")
    End If
    On Error GoTo 0
    '
    With PROXFECTMP
        On Error Resume Next
        .MoveLast
        If Err Then
          Call MENSERR(24, "No Existen Informes de Producción\en Rango de Fechas Seleccionado")
          On Error GoTo 0
          Screen.MousePointer = 1
        End If
        On Error GoTo 0
        FIN& = .RecordCount
        .MoveFirst
        '
          CODPARAM$(1) = "MEDIDA-3": DOCPARAM$(1) = VDEUV2$
          CODPARAM$(2) = "MEDIDA-2": DOCPARAM$(2) = PERIO$
          CAPARDOC% = 2
        '
          CODTABU1$(1) = "FECH-EMI"
          CODTABU1$(2) = "MEDIDA-1"
          CODTABU1$(3) = "COD-MAT"
          CODTABU1$(4) = "NOP-OPER"
          CODTABU1$(5) = "N-PLANO"
          CODTABU1$(6) = "DES-MAT"
          CODTABU1$(7) = "TIE-PREP"
          CODTABU1$(8) = "TIE-PROD"
          CODTABU1$(9) = "TIE-TOTA"
          CODTABU1$(10) = "TOT-TIEM"
          CODTABU1$(11) = "CANTIDAD"
          CODTABU1$(12) = "CAKILOS"
          CODTABU1$(13) = "REF-MAT1"
          CODTABU1$(14) = "REF-MAT2"
          CODTABU1$(15) = "REF-MAT3"
          '
          CACOLTAB1% = 15
        '
          Do While .EOF = False
          '
            EQUI$ = !Maquina
          '
            SEC = SECPERTE(EQUI$)
          '
            If SEC = XVALO(Mid(VDEVU1$, 1, 5)) Then
           '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = FECHATEX$(CVINT(!FeHorLib))
              TETABU1$(2, CAITAB1%) = !IdSubOr
              TETABU1$(3, CAITAB1%) = CODEXT$(!Cod_Inte)
              DOPE$ = TRIM$(Str$(!NumeOper))
              TETABU1$(4, CAITAB1%) = DOPE$
              '
              OrdeFabr.FindFirst "IdSubOr = '" & !IdSubOr & "' "
              If OrdeFabr.NoMatch = False Then
                On Error Resume Next
                TETABU1$(5, CAITAB1%) = OrdeFabr!Destino
                On Error GoTo 0
              End If
              '
              DESC$ = DESCRIT0$(!Cod_Inte)
              TETABU1$(6, CAITAB1%) = DESC$
              HHORA$ = HORATEX$(!Hora_Des)
              TETABU1$(7, CAITAB1%) = HHORA$
              HHORA1$ = HORATEX$(!Hora_Has)
              TETABU1$(8, CAITAB1%) = HHORA1$
              HHORA2$ = HORATEX$(!Hora_Fri)
              TETABU1$(9, CAITAB1%) = HHORA2$
              TIENETO = 60 * (DIFHORA(!Hora_Des, !Hora_Has) - DIFHORA(0, !Hora_Fri))
              If TIENETO < 0 Then
                 TIENETO = 0
              End If
              HHORA3$ = Left$(HORMINSE$(60 * TIENETO), 6)
              TETABU1$(10, CAITAB1%) = HHORA3$
              TETABU1$(11, CAITAB1%) = FORMATNUM$(!CantReal, "F5.0")
              TETABU1$(12, CAITAB1%) = FORMATNUM$(!CaNoConf, "F5.0")
              TETABU1$(13, CAITAB1%) = ""
              NOCON = 0
              If !CantReal > 0.01 Then
                 NOCON = (!CaNoConf / !CantReal) * 100
              End If
              TETABU1$(14, CAITAB1%) = FORMATNUM$(NOCON, "F5.1")
              CADEN = 0
              If TIENETO > 0 Then
                 TIEHOR = TIENETO / 60
                 CADEN = (!CantReal / TIEHOR)
              End If
              TETABU1$(15, CAITAB1%) = FORMATNUM$(CADEN, "F10.1")
           '
            End If
        .MoveNext
        Loop
        '
      End With
      '
      Call PRINTDOC("FAUTCONT")
      '
      '27/03/07 (FIN)
      '
End Sub


'*****/////COMANDO OCULTO/////*****
'Se oculto en el Picture6 ("Informes Parciales de Producción")
'el Command30("Informe de Tiempos de Operaciones Cumplidas)"
'*****/////**************/////*****
'
Private Sub ayd_Click()
Call Command29_Click
End Sub

Private Sub BdDOrFa_Click()
Call Command24_Click
End Sub

Private Sub CoCodPro_Click()
Call Command28_Click
End Sub

Private Sub CodProduc_Click()
Call Command3_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    HOII% = HOY(HOS$)
    VDEF$ = MKI$(COMESACT%) + MKI$(HOII%) + Space$(6)
10  OBX$ = OBJPLA$("DESHASFECHMAQ", VDEF$)
    If OBX$ = "" Then GoTo 99
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If HAS% < DES% Then GoTo 10
    MAQUISEL$ = TRIM$(Mid$(OBX$, 5, 6))
    If MAQUISEL$ = "" Then GoTo 10
    If ECOARCH$("PADMAQ", MAQUISEL$) = "" Then GoTo 10
    '
    FEC1 = CDbl(CVDAT(DES%))
    FEC2 = CDbl(CVDAT(HAS%))
    '
    Call ABREORFAB
    sqlx$ = " SELECT * FROM InfoFab "
    sqlx$ = sqlx$ + " WHERE  INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
    sqlx$ = sqlx$ + " AND INT(CDBL(FeHorLib))<= " & FEC2 & " "
    sqlx$ = sqlx$ + " AND Maquina = '" & MAQUISEL$ & "' "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    Set PROMAQTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    With PROMAQTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
        Call MENSERR(24, "No Existen Informes de Producción.")
        On Error GoTo 0
        GoTo 10
        End If
        On Error GoTo 0
        '
        JJ& = 0
        XX$ = REGBLAN$("INFOPEFA")
        Do While .EOF = False
          '
          ALTOP% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          On Error GoTo 0
          '
          LEGAJOP% = CVI(!LegaOps)
          Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
          Call REPLA(XX$, !IdSubOr, 3, 12)
          Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, !Maquina, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
          Call REPLA(XX$, !IdSubOr, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CantReal), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
            If !TiemEjec <= 0 Then
              RENDIM = 0
            Else
              RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
            End If
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
          '
          JJ& = JJ& + 1
          Call GRAREG("INFOPEFA", XX$, JJ&)
          .MoveNext
        Loop
    End With
    Call SETULTREG("INFOPEFA", JJ&)
    Call CIERRARCH("INFOPEFA")
    '
    Call HEADERS("INFOPEFA", "")
    Call HEADERS("INFOPEFA", "Equipo: " + TRIM$(MAQUISEL$) + " - " + TRIM$(ECOARCH$("PADMAQ", MAQUISEL$)))
    Call CONECRUN("*INFAMAQ", "Informes de Producción por Equipo")
    Screen.MousePointer = 1
    GoTo 10
    '
99  Command15.Enabled = True
    Call FRESHFORM(Caption)
    '
End Sub

Sub Command19_Click()
   Command19.Enabled = False
   OPXX% = ALTERNA%("Padron de Operarios\Células de Producción\Integrantes de Células")
   If OPXX% = 1 Then
        Call FINAL("*PADRON")
        AppActivate Caption
        SendKeys Chr$(1)
      ElseIf OPXX% = 2 Then
        Call DEFCELULAS
      ElseIf OPXX% = 3 Then
        Call CARCELULAS
   End If
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   TTXX$ = "Los Listados Presentan las Operaciones Pendientes"
   TTXX$ = TTXX$ + "\para las Ordenes de Fabricación en Curso de Ejecución"
   TTXX$ = TTXX$ + "\al Momento del Último Cálculo de Secuenciación CRP."
   TTXX$ = TTXX$ + "\   \Opciones de Agrupamiento del Listado:"
   TTXX$ = TTXX$ + "\\Por Número de O/F\Por Código de Producto"
   OPXX% = COMALTER%(TTXX$)
   '
   If OPXX% = 1 Then
        Call GENEOPERFAP
        Call GESALPROP(1)
      ElseIf OPXX% = 2 Then
        Call GENEOPERFAP
        Call GESALPROP(2)
   End If
   '
   Command2.Enabled = True
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Call CONECRUN("ORFABR04", "Ordenes de Fabricación")
    Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    ACONEC$ = "CARMAQ04"
    Call CONECRUN(ACONEC$, "")
    Command25.Enabled = True
End Sub

Sub Command26_Click()
    Command26.Enabled = False
    Dim LEGAI%(2048), TIEREAL(2048), TIEMSTA(2048), RENDPROM(2048)
    Dim REGMIN&(2048), VAMIN(2048), REGMAX&(2048), VAMAX(2048)
    '
    CANOPERS% = 0
    RENDIMA = XVALO(CONTROL("PADIAMOD", "RENDIMAX"))
    If RENDIMA < 1 Then RENDIMA = 99999
    '
    'ACA AGREGO
    ALT% = 0
    '27/03/07 (OT-07-0087)
    ALTMAX% = 3
    OPC$ = "Listados Generales\Por Operario\Por Pieza y Operación"
    If CONTROL$("IFABRICA", "FAUTCONT") <> "" Then
       OPC$ = OPC$ + "\Listado de AutoControl"
       ALTMAX% = 4
    End If
    '27/03/07 (FIN)
    '
5   RTA1% = ALTERNA%(OPC$)
      If RTA1% = 1 Then
          GoTo 10
        ElseIf RTA1% = 2 Then
          Call CONPORLEGA
          GoTo 5
        ElseIf RTA1% = 3 Then
          INFOPIE.Show 1
          GoTo 5
        '27/03/07 (OT-07-0087)
        ElseIf RTA1% = 4 Then
          If RTA1% <= ALTMAX% Then Call AUTOCONTROL
          GoTo 99
        '27/03/07 (FIN)
        Else
          GoTo 99
      End If
    
10  Call DESHASFECHA(DES1%, HAS1%, PERIO$)
    VDEVU$ = PERIO$
    If VDEVU$ <> "" Then
      FF% = FECHANUM(VDEVU$)
      SORTLIST.Clear
      Call BLOQARCH("INFOPEFA")
      JJ& = 0
      '
      Screen.MousePointer = 11
      FEC1 = CDbl(CVDAT(DES1%))
      FEC2 = CDbl(CVDAT(HAS1%))
      '
      Call ABREORFAB
      sqlx$ = " SELECT * FROM InfoFab "
      sqlx$ = sqlx$ + " WHERE  INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
      sqlx$ = sqlx$ + " AND  INT(CDBL(FeHorLib))  <=  " & FEC2 & " "
      sqlx$ = sqlx$ + " ORDER BY IdSubOr "
      On Error Resume Next
      Set PROXFECTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Registros no Procesables.\Debe Ejecutar Opción de Control")
        Call Command32_Click
        GoTo 99
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      Call COPYSTRU("INFOPEFA", "INFOPEMA")
      With PROXFECTMP
        On Error Resume Next
        .MoveLast
        If Err Then
          Call MENSERR(24, "No Existen Informes de Producción\en Rango de Fechas Seleccionado")
          On Error GoTo 0
          Screen.MousePointer = 1
          GoTo 10
        End If
        On Error GoTo 0
        FIN& = .RecordCount
        .MoveFirst
        '
        JJ& = 0: KK& = 0: JTL& = 0: HUBOATMUL% = 0
        XX$ = REGBLAN$("INFOPEFA")
        Do While .EOF = False
          '
          JTL& = JTL& + 1
          Call TRACE(20, JTL&, FIN&)
          '
          ALTOP% = 0
          ATMUL% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          ATMUL% = !AtMulti
          On Error GoTo 0
          '
          If ATMUL% > 0 Then
             HUBOATMUL% = 1
             GoTo 29
          End If
          '
          Call REPLA(XX$, !IdSubOr, 3, 12)
          Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
          Call REPLA(XX$, DOPE$, 17, 42)
          On Error Resume Next
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 59, 2)
          Call REPLA(XX$, MKI$(!Hora_Des), 61, 2)
          Call REPLA(XX$, MKI$(!Hora_Has), 63, 2)
          On Error GoTo 0
          Call REPLA(XX$, !Maquina, 65, 6)
            HHORA$ = Format(!FeHorLib, "hh:mm:ss")
            If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then
               On Error Resume Next
               HHORA$ = HORATEXNOR$(!Hora_Has)
               On Error GoTo 0
            End If
            FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
            If IsNull(!Hora_Has) = True Then
              Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 63, 2)
            End If
          Call REPLA(XX$, !IdSubOr, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          '
          'Se suma Cant. Real y Cant. No Conf
          NoConf# = XVALO(!CaNoConf)
          Real# = XVALO(!CantReal)
          TotCant# = NoConf# + Real#
          Call REPLA(XX$, MKD$(TotCant#), 139, 8)
          '
          'Call REPLA(XX$, MKD$(!CantReal), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf, ALTOP%)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
          '
          ' ENJUAGUE ESCORIAL
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "ESCORIAL") > 0 Then
             'TSTANBA = TIESTAN(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
             TSTANBA = TIEMAQUI(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
             TSTAN = 60 * TIEMAQUI(!Cod_Inte, !NumeOper, TotCant#, !NuOrInf, ALTOP%)
             CADEN = 0
             If TSTANBA > 0 Then CADEN = 60 / TSTANBA
             Call REPLA(XX$, MKS$(CADEN), 155, 4)
          End If
            '
            If !TiemEjec <= 0 Then
                RENDIM = 0
              Else
                RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
            End If
            ' If RENDIM > RENDIMA Then RENDIM = RENDIMA
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          '
          LEGACELU$ = !LegaOps
          For KKI% = 1 To Len(LEGACELU$) Step 2
             LEGAJOP% = CVI(Mid$(LEGACELU$, KKI%, 2))
             If LEGAJOP% > 0 Then
                Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
                Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
                '
                JJ& = JJ& + 1
                Call GRAREG("INFOPEFA", XX$, JJ&)
                If KKI% = 1 Then
                  KK& = KK& + 1
                  Call GRAREG("INFOPEMA", XX$, KK&)
                End If
                '
                For KKII% = 1 To CANOPERS%
                  If LEGAI%(KKII%) = LEGAJOP% Then
                    TIEREAL(KKII%) = TIEREAL(KKII%) + !TiemEjec * 60
                    TIEMSTA(KKII%) = TIEMSTA(KKII%) + TSTAN
                    If RENDIM > 0.05 Then
                      If RENDIM < VAMIN(KKII%) Then
                        VAMIN(KKII%) = RENDIM
                        REGMIN&(KKII%) = JJ&
                      End If
                      If RENDIM > VAMAX(KKII%) Then
                        VAMAX(KKII%) = RENDIM
                        REGMAX&(KKII%) = JJ&
                      End If
                    End If
                    GoTo 26
                  End If
                Next KKII%
                CANOPERS% = CANOPERS% + 1
                LEGAI%(CANOPERS%) = LEGAJOP%
                TIEREAL(CANOPERS%) = !TiemEjec * 60
                TIEMSTA(CANOPERS%) = TSTAN
                VAMIN(CANOPERS%) = RENDIM
                REGMIN&(CANOPERS%) = JJ&
                VAMAX(CANOPERS%) = RENDIM
                REGMAX&(CANOPERS%) = JJ&
             End If
26        Next KKI%
          '
29      .MoveNext
        Loop
      End With
      '
      Call SETULTREG("INFOPEFA", JJ&)
      Call CIERRARCH("INFOPEFA")
      Call SETULTREG("INFOPEMA", KK&)
      Call CIERRARCH("INFOPEFA")
      '
      REBLA$ = REGBLAN$("RESUREND")
      For KKL& = 1 To CANOPERS%
        XX$ = REBLA$
        Call REPLA(XX$, MKI$(LEGAI%(KKL&)), 1, 2)
        Call REPLA(XX$, MKS$(TIEREAL(KKL&)), 3, 4)
        Call REPLA(XX$, MKS$(TIEMSTA(KKL&)), 7, 4)
        '
        OBSEX$ = ""
        If TIEREAL(KKL&) <= 0 Then
             RENDIM = 0
           Else
             RENDIM = (100 * TIEMSTA(KKL&)) / TIEREAL(KKL&)
        End If
        If RENDIM > RENDIMA Then
          RENDIM = RENDIMA
          OBSEX$ = "Ajustado a " + TRIM$(Str$(RENDIMA))
        End If
        Call REPLA(XX$, MKS$(RENDIM), 11, 4)
        Call REPLA(XX$, OBSEX$, 15, 34)
        Call GRAREG("RESUREND", XX$, KKL&)
        RENDPROM(KKL&) = RENDIM
        If REGMIN&(KKL&) > 0 And REGMIN&(KKL&) <= JJ& Then
          XX$ = REGLEIDO$("INFOPEFA", REGMIN&(KKL&))
          Call REPLA(XX$, "min", 169, 6)
          Call GRAREG("INFOPEFA", XX$, REGMIN&(KKL&))
        End If
        If REGMAX&(KKL&) > 0 And REGMAX&(KKL&) <= JJ& Then
          XX$ = REGLEIDO$("INFOPEFA", REGMAX&(KKL&))
          Call REPLA(XX$, "MAX", 169, 6)
          Call GRAREG("INFOPEFA", XX$, REGMAX&(KKL&))
        End If
      Next KKL&
      KKL& = CANOPERS%
      Call SETULTREG("RESUREND", KKL&)
      '
      Call COPYSTRU("PADRON", "PADREND")
      FIN& = ULTREG&("PADRON")
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        XX$ = REGLEIDO$("PADRON", U&)
        LEGAJ% = CVI(Left$(XX$, 2))
        RENDX = 0
        For KKL& = 1 To CANOPERS%
          If LEGAI%(KKL&) = LEGAJ% Then
            RENDX = RENDPROM(KKL&)
            GoTo 39
          End If
        Next KKL&
39      APNO$ = TRIM$(Mid$(XX$, 3, 30))
        If RENDX >= 0.05 Then
          APNO$ = Left$(APNO$, 24) + " (" + TRIM$(FORMATNUM$(RENDX, "F5.1")) + ")"
        End If
        Call REPLA(XX$, APNO$, 3, 30)
        Call GRAREG("PADREND", XX$, U&)
      Next U&
      Call SETULTREG("PADREND", FIN&)
      Call CIERRARCH("PADREND")
      '
      FEX = FF%
      Call HEADERS("INFOPEFA", "")
      Call HEADERS("INFOPEFA", "Corresponde a: " + PERIO$)
      Call HEADERS("INFOPEMA", "")
      Call HEADERS("INFOPEMA", "Corresponde a: " + PERIO$)
      '
      Call HEADERS("RESUREND", "")
      Call HEADERS("RESUREND", "Corresponde a: " + PERIO$)
      Screen.MousePointer = 1
      '
      If HUBOATMUL% > 0 Then
         Call COMUNI("Los Informes Correspondientes a Atención Múltiple\Se Suprimieron del Presente Listado.")
      End If
      '
      OPX% = COMALTER%("Elija Agrupamiento:\\Por Operario\Por Maquina")
      Call FILESORT("INFOPEFA", "", 6, 8, "ASC")
      If OPX% = 1 Then
           OPY% = COMALTER%("Formato de Salida:\\Detalle\Resumen")
           If OPY% = 1 Then
                Call FILESORT("INFOPEFA", "", 1, 2, "ASC")
                Call CONECRUN("*INFAOPE", "Informes de Producción por Operario")
              ElseIf OPY% = 2 Then
                Call FILESORT("RESUREND", "", 1, 2, "ASC")
                Call CONECRUN("*RESUREN", "Resumen de Horas y Rendimiento por Operario")
           End If
         ElseIf OPX% = 2 Then
           Call FILESORT("INFOPEMA", "", 9, 12, "ASC")
           Call CONECRUN("*INFADIA", "Informes de Producción por Fecha")
      End If
      Call CIERRARCH("INFOPEFA")
      Screen.MousePointer = 1
      ' GoTo 10
      '
    End If
99  Command26.Enabled = True
    Call FRESHFORM(Caption)
    Call FRESHALL
    '
End Sub

Sub CONPORLEGA()
    '
    Dim LEGAI%(2048), TIEREAL(2048), TIEMSTA(2048)
    '
    Call COPYSTRU("PADRON", "PADREND")
    CANOPERS% = 0
    RENDIMA = XVALO(CONTROL("PADIAMOD", "RENDIMAX"))
    If RENDIMA < 1 Then RENDIMA = 99999
    HOII% = HOY(HOS$)
    VDEF$ = MKI$(0) + MKI$(COMESACT%) + MKI$(HOII%)
    '
10  VDEVU$ = OBJPLA$("SELOPERFECH", VDEF$)
    If VDEVU$ = "" Then GoTo 99
    VDEF$ = VDEVU$
    '
    LEGASEL% = CVI(Mid$(VDEVU$, 1, 2))
    DES1% = CVI(Mid$(VDEVU$, 3, 2))
    HAS1% = CVI(Mid$(VDEVU$, 5, 2))
    '
    'VALIDACIONES
    If ECOARCH$("PADRON", MKI$(LEGASEL%)) = "" Then
     Call MENSERR(24, "Legajo Inexistente o no Válido")
     GoTo 10
    End If
    '
    If DES1% < 33 Or HAS1% < DES1% Then
      Call MENSERR(24, "Rango de Fechas no Válido")
      GoTo 10
    End If
    'SI ALGO NO VALIDA A 10
    '
    FF% = FECHANUM(VDEVU$)
    SORTLIST.Clear
    Call BLOQARCH("INFOPEFA")
    JJ& = 0
    '
    Screen.MousePointer = 11
    FEC1 = CDbl(CVDAT(DES1%))
    FEC2 = CDbl(CVDAT(HAS1%))
    '
    Call ABREORFAB
    sqlx$ = " SELECT * FROM InfoFab "
    sqlx$ = sqlx$ + " WHERE  INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
    sqlx$ = sqlx$ + " AND  INT(CDBL(FeHorLib))  <=  " & FEC2 & " "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    '
    Set PROXFECTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    '
    With PROXFECTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No Existen Informes de Producción\en Rango de Fechas Seleccionado")
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 10
       End If
       On Error GoTo 0
       '
       JJ& = 0
       REMIN& = 0: VAMIN = 999999
       REMAX& = 0: VAMAX = 0
       SUSTAN = 0: SUEJEC = 0
       XX$ = REGBLAN$("INFOPEFA")
       Do While .EOF = False
         '
         ALTOP% = 0
         ATMUL% = 0
         On Error Resume Next
         ALTOP% = !AltOper
         ATMUL% = !AtMulti
         On Error GoTo 0
         '
         LEGACELU$ = !LegaOps
         For KKI% = 1 To Len(LEGACELU$) Step 2
            LEGAJOP% = CVI(Mid$(LEGACELU$, KKI%, 2))
            If LEGAJOP% = LEGASEL% Then
              Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
              Call REPLA(XX$, !IdSubOr, 3, 12)
              Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
                DOPE$ = TRIM$(Str$(!NumeOper))
                If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
                DOPE$ = TRIM$(AJUSTI$(DOPE$, 7) + !DescOper)
                If ATMUL% > 0 Then DOPE$ = DOPE$ + "-AM"
              Call REPLA(XX$, DOPE$, 17, 42)
              On Error Resume Next
              Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 59, 2)
              Call REPLA(XX$, MKI$(!Hora_Des), 61, 2)
              Call REPLA(XX$, MKI$(!Hora_Has), 63, 2)
              On Error GoTo 0
              Call REPLA(XX$, !Maquina, 65, 6)
                HHORA$ = Format(!FeHorLib, "hh:mm:ss")
                If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
                FFECH$ = Format(!FeHorLib, "dd/mm/yy")
              Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
              Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
                If IsNull(!Hora_Has) = True Then
                  Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 63, 2)
                End If
              Call REPLA(XX$, !IdSubOr, 75, 12)
              Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
              Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
              Call REPLA(XX$, DOPE$, 91, 48)
              'Se suma Cant. Real y Cant. No Conf
              NoConf# = XVALO(!CaNoConf)
              Real# = XVALO(!CantReal)
              TotCant# = NoConf# + Real#
              Call REPLA(XX$, MKD$(TotCant#), 139, 8)
              '
              Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
              TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf, ALTOP%)
              Call REPLA(XX$, MKS$(TSTAN), 155, 4)
              '
              ' ENJUAGUE ESCORIAL
              EPAC$ = TRIM$(UCase$(EMPREAC$))
              If InStr(EPAC$, "ESCORIAL") > 0 Then
                 'TSTANBA = TIESTAN(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
                 TSTANBA = TIEMAQUI(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
                 TSTAN = 60 * TIEMAQUI(!Cod_Inte, !NumeOper, TotCant#, !NuOrInf, ALTOP%)
                 CADEN = 0
                 If TSTANBA > 0 Then CADEN = 60 / TSTANBA
                 Call REPLA(XX$, MKS$(CADEN), 155, 4)
              End If
                '
                If !TiemEjec <= 0 Then
                    RENDIM = 0
                  Else
                    RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
                End If
              Call REPLA(XX$, MKS$(RENDIM), 163, 4)
              Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
              '
              If ATMUL% > 0 Then
                 Call REPLA(XX$, "AM", 169, 6)
              End If
              '
              JJ& = JJ& + 1
              Call GRAREG("INFOPEFA", XX$, JJ&)
              '
              If ATMUL% < 1 Then
                SUSTAN = SUSTAN + TSTAN
                SUEJEC = SUEJEC + !TiemEjec
              End If
              '
              If RENDIM < VAMIN Then
                 VAMIN = RENDIM
                 REMIN& = JJ&
              End If
              If RENDIM > VAMAX Then
                 VAMAX = RENDIM
                 REMAX& = JJ&
              End If
              '
            End If
         Next KKI%
         '
29       .MoveNext
       Loop
    End With
    '
    If JJ& > 0 Then
      If REMIN& > 0 And REMIN& <= JJ& Then
        XX$ = REGLEIDO$("INFOPEFA", REMIN&)
        Call REPLA(XX$, "min", 169, 6)
        Call GRAREG("INFOPEFA", XX$, REMIN&)
      End If
      If REMAX& > 0 And REMAX& <= JJ& Then
        XX$ = REGLEIDO$("INFOPEFA", REMAX&)
        Call REPLA(XX$, "MAX", 169, 6)
        Call GRAREG("INFOPEFA", XX$, REMAX&)
      End If
    End If
    '
    Call SETULTREG("INFOPEFA", JJ&)
    Call CIERRARCH("INFOPEFA")
    '
    VAPROM = 9999999
    If SUEJEC >= 0.05 Then
       VAPROM = 100 * SUSTAN / (SUEJEC * 60)
    End If
    TTYY$ = "  ( " + TRIM$(FORMATNUM$(VAMIN, "F8.1"))
    TTYY$ = TTYY$ + " / " + TRIM$(FORMATNUM$(VAMAX, "F8.1"))
    TTYY$ = TTYY$ + " / " + TRIM$(FORMATNUM$(VAPROM, "F8.1")) + " )"
    '
    XX$ = FILTERGETRECORD("PADRON", 1, MKI$(LEGASEL%))
    APNO$ = TRIM$(Mid$(XX$, 3, 30))
    If VAPROM >= 0.05 Then
        APNO$ = Left$(APNO$, 24) + " (" + TRIM$(FORMATNUM$(VAPROM, "F5.1")) + ")"
    End If
    Call REPLA(XX$, APNO$, 3, 30)
    Call FILTERSETRECORD("PADREND", 1, MKI$(LEGASEL%), XX$)
    '
    FEX = FF%
    Call HEADERS("INFOPEFA", "")
    Call HEADERS("INFOPEFA", "Corresponde a: " + TRIM$(ECOARCH$("PADRON", MKI$(LEGASEL%))) + TTYY$)
    '
    Call FILESORT("INFOPEFA", "", 6, 8, "ASC")
    Call FILESORT("INFOPEFA", "", 1, 2, "ASC")
    Call CONECRUN("*INFAOPE", "Informes de Producción por Operario")
    Screen.MousePointer = 1
    Call CIERRARCH("INFOPEFA")
    GoTo 10
     '
99  End Sub
'
Private Sub Command27_Click()
    Command27.Enabled = False
    'NUEVO
      Call ABREORFAB
    '
    sqlx$ = " SELECT * FROM ORDEFABR "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveLast
      If Err Then
         Call MENSERR(24, "NO Existen Ordenes de Fabricación. Imposible Continuar la Carga.")
         On Error GoTo 0
         GoTo 99
      End If
      .MoveFirst
      FIN& = .RecordCount
      JJ& = 0: JJ1& = 0: IDANT$ = ""
      Do While (.EOF = False)
        JJ1& = JJ1& + 1
        Call TRACE(20, JJ1&, FIN&)
        If !IdSubOr <> IDANT$ Then
           Y$ = REGBLAN("ECORDEF")
           '
           Call REPLA(Y$, !IdSubOr, 1, 12)
           Call REPLA(Y$, !Referen, 13, 52)
           '
           JJ& = JJ& + 1
           Call GRAREG("ECORDEF", Y$, JJ&)
           IDANT$ = !IdSubOr
        End If
        '
      .MoveNext
      Loop
    End With
    Call SETULTREG("ECORDEF", JJ&)
    Call CIERRARCH("ECORDEF")
    'HASTA ACA NUEVO
10  Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación.")
    NUORDFAB$ = TRIM$(VALACT1$("ECORDEF"))
    '
11  NUORDFAB$ = AJUSTI$(VALACT1$("ECORDEF"), 12)
    If TRIM$(NUORDFAB$) = "" Then
       Command27.Enabled = True
       Call FRESHFORM(Caption)
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    Call ABREORFAB
    sqlx$ = " SELECT * FROM InfoFab "
    sqlx$ = sqlx$ + " WHERE  IdSubOr  =  '" & NUORDFAB$ & "' "
    sqlx$ = sqlx$ + " ORDER BY NumeOper "
    Set PROXFECTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    With PROXFECTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Informes de Producción.")
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 10
      End If
      On Error GoTo 0
      '
      JJ& = 0
      XX$ = REGBLAN$("INFOPEFA")
      Do While .EOF = False
          '
          ALTOP% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          On Error GoTo 0
          '
          LEGAJOP% = CVI(!LegaOps)
          Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
          Call REPLA(XX$, !IdSubOr, 3, 12)
          Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, !Maquina, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
          Call REPLA(XX$, !IdSubOr, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CantReal), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
            If !TiemEjec <= 0 Then
              RENDIM = 0
            Else
              RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
            End If
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
          '
          JJ& = JJ& + 1
          Call GRAREG("INFOPEFA", XX$, JJ&)
          .MoveNext
      Loop
    End With
    '
    Call SETULTREG("INFOPEFA", JJ&)
    Call CIERRARCH("INFOPEFA")
    '
    Call FINAL("*COINFOF1")
99  Screen.MousePointer = 1
    Command27.Enabled = True
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    VDEF$ = MKI$(COMESACT%) + MKI$(HOII%) + Space$(6)
10  OBX$ = OBJPLA$("DESHASFECHCOD", VDEF$)
    If OBX$ = "" Then GoTo 99
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If HAS% < DES% Then GoTo 10
    CIIX% = CVI(Mid$(OBX$, 5, 2))
    If CIIX% < 1 Then GoTo 10
    FEC1 = CDbl(CVDAT(DES%))
    FEC2 = CDbl(CVDAT(HAS%))
    '
    Call ABREORFAB
    sqlx$ = " SELECT * FROM InfoFab"
    sqlx$ = sqlx$ + " WHERE INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
    sqlx$ = sqlx$ + " AND INT(CDBL(FeHorLib))<= " & FEC2 & " "
    sqlx$ = sqlx$ + " AND CINT(Cod_Inte) = '" & CIIX% & "' "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    Set PROMAQTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    With PROMAQTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
        Call MENSERR(24, "No Existen Informes de Producción.")
        On Error GoTo 0
        GoTo 10
        End If
        On Error GoTo 0
        '
        JJ& = 0
        XX$ = REGBLAN$("INFOPEFA")
        Do While .EOF = False
          '
          ALTOP% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          On Error GoTo 0
          '
          LEGAJOP% = CVI(!LegaOps)
          Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
          Call REPLA(XX$, !IdSubOr, 3, 12)
          Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
          Call REPLA(XX$, !DescOper, 17, 48)
          Call REPLA(XX$, !Maquina, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
          Call REPLA(XX$, !IdSubOr, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CantReal), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
            If !TiemEjec <= 0 Then
              RENDIM = 0
            Else
              RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
            End If
          Call REPLA(XX$, MKS$(RENDIM), 163, 4)
          Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
          '
          JJ& = JJ& + 1
          Call GRAREG("INFOPEFA", XX$, JJ&)
          .MoveNext
        Loop
    End With
    '
    Call SETULTREG("INFOPEFA", JJ&)
    Call CIERRARCH("INFOPEFA")
    '
    Call FILESORT("INFOPEFA", "", 3, 4, "ASC")
    '
    Call HEADERS("INFOPEFA", "")
    Call HEADERS("INFOPEFA", "Codigo: " + TRIM$(CODEXT$(CIIX%)) + " - " + TRIM$(DESCRIT0$(CIIX%)))
    Call CONECRUN("*INFACOD", "Informes de Producción por Articulo")
    Screen.MousePointer = 1
    GoTo 10
     '
99  Command28.Enabled = True
    Call FRESHFORM(Caption)
    '
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("INFOFAB")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   TTXX$ = "Los Listados Presentan las Operaciones Informadas"
   TTXX$ = TTXX$ + "\para las Ordenes de Fabricación en Curso de Ejecución."
   TTXX$ = TTXX$ + "\   \Opciones de Detalle de Presentación:"
   TTXX$ = TTXX$ + "\\Detalle por Informe\Acumulado por Operación\Saldos de Productos en Proceso"
   OPXX% = COMALTER%(TTXX$)
   '
   If OPXX% = 1 Then
        Call CALCUPROCE0
        Call CONECRUN("*ACUPROCE", "Listado acumulado de procesos")
      ElseIf OPXX% = 2 Then
        Call CALCUPROCE0
        Call CALCUPROCE1
        Call CONECRUN("*ACUPROC1", "Listado de Totales")
      ElseIf OPXX% = 3 Then
        Call CALCUPROCE0
        Call CALCUPROCE1
        Call CALCUPROCE2
        Call CONECRUN("*ACUPROC2", "Saldos de Productos en Proceso")
   End If
   Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   Command30.Enabled = False
   'If DERACCE%("MODINFOP", "Modificar Partes de Fabricacion") > 0 Then
   '   Call INFOPER
   'End If
   Command30.Enabled = True
End Sub

Private Sub Command31_Click()
    '
    Command31.Enabled = False
    Call CONECRUN("EFIMOB04", "Control de Productividad de Mano de Obra")
    Command31.Enabled = True
    '
End Sub

Private Sub Command32_Click()
    Command32.Enabled = False
    Call ABREORFAB
    Screen.MousePointer = 11
    With InfoFab
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        FIN& = .RecordCount
        JJ& = 0
        RECU% = 0
        .MoveFirst
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          If IsNull(!FeHorLib) = True Then
            .Edit
            !FeHorLib = CVDAT(Int(FECHANUM("01/01/80")))
            .Update
            RECU% = 1
          End If
        .MoveNext
        Loop
      End If
    End With
    Screen.MousePointer = 1
    '
    TTXX$ = "Control Completo"
    If RECU% > 0 Then TTXX$ = TTXX$ + "\con Registros Recuperados."
    Call COMUNI(TTXX$)
    '
    Command32.Enabled = True
Exit Sub
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
   Command33.Enabled = False
   AJUPROCE04.Show 1        ' AJUSTE DE SALDOS PRODUCTOS EN PROCESO
   Command33.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call GENEOPERFAP
    Call FINAL("*SECAROF")
    Command4.Enabled = True
End Sub
'

Private Sub Command5_Click()
    Command4.Enabled = False
    Call GENEOPERFAP        'NUEVO
    Call FINAL("*SECAREQ")
    Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
'    Call MENSERR(24, "Informe Diario No Disponible\por el Momento.\ \Debe Informar por Número\de Orden de Fabricación")
'    Call INFAGEN
'    'Call INFADIA
    '\\\OT-06-0252-WAR-26/06/06///
    If EXISTE%("APROYE04.EXE") < 1 Then
      Call MENSERR(24, "Modulo APROYE04 no Instalado\en la Presente Versión.")
    Else
      RTA1% = ALTERNA%("Registrar Informe de Proyecto\Anular Informe de Proyecto")
      If RTA1% = 1 Then
        Call INFAPROY
      ElseIf RTA1% = 2 Then
        Call ANUINFAPROY
      End If
    End If
    '\\\OT-06-0252-WAR-FIN///
    '
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    '\\\OT-06-0309-WAR-06/07/06///
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ESCORIAL") > 0 Or InStr(EPAC$, "ALEAR") Then
        PADIAMOD04.Show 1
      Else
        Call INFAGEN
    End If
    Command7.Enabled = True
End Sub



Private Sub Command9_Click()
   Command9.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMASTO"
   Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
   Command9.Enabled = True
End Sub

Private Sub CoNumOF_Click()
Call Command27_Click
End Sub

Private Sub CoProFec_Click()
Call Command26_Click
End Sub

Private Sub CtrlOpPen_Click()
Call Command32_Click
End Sub

Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("EFIMOB04.EXE") < 1 Then
      Command31.Enabled = False
      Label11.Enabled = False
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
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREORFAB
     sqlx$ = " SELECT * FROM OPERFAB "
     sqlx$ = sqlx$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     sqlx$ = sqlx$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     sqlx$ = sqlx$ + " ORDER BY NumeOper "
     '
     Set FABRTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
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
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
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
            NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
            Call REPLA(YY$, NOPX$, 3, 58)
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
          End If
       .MoveNext
       Loop
     End With
        
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
       GoTo 1405
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     NOPE% = Val(Mid$(XX$, 3, 5))
     SALDO# = Val(Mid$(XX$, 49, 12))
     Call CARGAINF(NROOF$, NOPE%)
     If NOPE% < 1 Then
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

Sub INFADIA()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "INFORMES DE PRODUCCION - GENERALES"
     '
100  VDEF$ = MKI$(HO%) + Space$(6)
101  DIAEQUI$ = OBJPLA$("DIAEQUI", VDEF$)
     If DIAEQUI$ = "" Then Exit Sub
     '
     DIA% = CVI(Left$(DIAEQUI$, 2))
     If DIA% > HO% Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 100
     End If
     '
     MAQUI$ = Mid$(DIAEQUI$, 3, 6)
     If TRIM$(ECOARCH("PADMAQ", MAQUI$)) = "" Then
         Call MENSERR(24, "Equipo no Definido")
         GoTo 100
     End If
     '
     Screen.MousePointer = 11
     RGFSX$ = RECFILT$("OPERFAP", 7, MAQUI$)
     If Len(RGFSX$) < 4 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para el Equipo Elegido.")
         Screen.MousePointer = 1
         GoTo 100
     End If
     '
     J% = 0
     MENSAIN$ = "Inconsistencia en Base de Datos de Ope-\raciones de Fabricación.\  \Los Punteros de  Operaciones Pendientes\no se Corresponden con los Registros de\Operaciones Programadas.\  \Consulte con su Soporte de Sistemas."
     Call FRESHECHO("!OPERINF")
     For K% = 1 To Len(RGFSX$) Step 4
        REGI& = CVS(Mid$(RGFSX$, K%, 4))
        XX$ = REGLEIDO$("OPERFAP", REGI&)
        DES% = CVI(Mid$(XX$, 135, 2))
        HAS% = CVI(Mid$(XX$, 139, 2))
        If DIA% >= DES% Then
           If DIA% <= HAS% Then
              CAPROD# = CVD(Mid$(XX$, 79, 8))
              CAREA# = CVD(Mid$(XX$, 87, 8))
              SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
              Call REPLA(XX$, MKD$(SALDO#), 95, 8)
              Call GRAREG("OPERFAP", XX$, REGI&)
              '
              INTEN% = 0
110           REGO& = CVS(Mid$(XX$, 221, 4))
              If REGO& < 1 Or REGO& > ULTREG&("OPERFAB") Then
                 Call MENSERR(24, MENSAIN$)
                 Screen.MousePointer = 1
                 Exit Sub
              End If
              '
              ZZ$ = REGLEIDO$("OPERFAB", REGO&)
              If Left$(ZZ$, 86) <> Left$(XX$, 86) Then
If HOY(HOS$) > FECHANUM("10/06/01") Then GoTo 111
                If INTEN% = 0 Then
                  Screen.MousePointer = 11
                  For KKL& = 1 To ULTREG&("OPERFAB")
                    If Left$(REGLEIDO$("OPERFAB", KKL&), 86) = Left$(XX$, 86) Then
                       XX$ = REGLEIDO$("OPERFAP", REGI&)
                       Call REPLA(XX$, MKS$(KKL&), 221, 4)
                       Call GRAREG("OPERFAP", XX$, REGI&)
                       Screen.MousePointer = 1
                       INTEN% = INTEN% + 1
                       GoTo 110
                    End If
                  Next KKL&
                  Screen.MousePointer = 1
                End If
111             Call MENSERR(24, MENSAIN$)
                Screen.MousePointer = 1
                Exit Sub
              End If
              '
              If SALDO# > 0.05 Then
                 NROOF$ = Left$(XX$, 12)
                 CI% = CVI(Mid$(XX$, 13, 2))
                 NOPE% = CVI(Mid$(XX$, 15, 2))
                 DOPE$ = Mid$(XX$, 17, 48)
                 DESTI$ = Mid$(XX$, 65, 8)
                 CAINF# = SALDO#
                 FEOP% = CVI(Mid$(XX$, 139, 2))
                 HOOP% = CVI(Mid$(XX$, 141, 2))
                 MALIB$ = "S"
                 TITOT = CVS(Mid$(XX$, 131, 4))
                 '
                 J% = J% + 1
                 YY$ = MKI$(J%) + Space$(58) + MKS$(REGI&)
                 NOPX$ = AJUSTI$(NROOF$, 13) + AJUSTI$(DESTI$, 9)
                 NOPX$ = NOPX$ + AJUSTD$(Str$(NOPE%), 5) + " - " + AJUSTI$(DOPE$, 38)
                 NOPX$ = Left$(NOPX$, 46)
                 NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
                 Call REPLA(YY$, NOPX$, 3, 58)
                 JJ& = J%
                 Call GRAREG("!OPERINF", YY$, JJ&)
                 '
              End If
           End If
        End If
     Next K%
     '
     If J% < 1 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Fecha y ese Equipo.")
         Screen.MousePointer = 1
         GoTo 100
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
200  FEX = DIA%
     Call SELECHO("!OPERINF/Operaciones Programadas para el " + FECHATEX$(FEX) + " - " + TRIM$(MAQUI$))
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
         GoTo 100
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     REGI& = CVS(Mid$(XX$, 61, 4))
     Call CARGAINF(NROOF$, NOPE%)
     If NOPE% < 1 Then
          GoTo 200
        Else
          VDEF$ = DIAEQUI$
          GoTo 101
     End If
     '
End Sub
'
Sub INFOPER()      ' INFORMES DE OPERACION POR O/F
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "INFORMES DE OPERACION POR O/F"

100  NROOF$ = ""

     Call FRESHECHO("ECORDEF")
     If ULTREG&("ECORDEF") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación en Curso.")
         Exit Sub
     End If
     '
1405 Call SELECHO("ECORDEF/Ordenes de Trabajo en Curso")
     Call FRESHALL
     '
1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEF"), 12)
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
     RFF$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
     JJ& = 0
     For KK% = 1 To Len(RFF$) Step 4
       REOPER& = CVS(Mid$(RFF$, KK%, 4))
       ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
       NORFA$ = Left$(ZZ$, 12)
       If NORFA$ = NROOF$ Then
         CAREA# = CVD(Mid$(ZZ$, 87, 8))
         If Mid$(ZZ$, 177, 6) <> String$(6, 0) Then ' solo las ya informadas
            MAQUI$ = Mid$(ZZ$, 151, 6): If TRIM$(MAQUI$) = "" Then MAQUI$ = Mid$(ZZ$, 73, 6)
            FERE% = CVI(Mid$(ZZ$, 157, 2))
            HORE% = CVI(Mid$(ZZ$, 159, 2))
            LEGOPER$ = Mid$(ZZ$, 165, 6)
            LEGA4% = CVI(Mid$(ZZ$, 171, 2))
            If LEGA4% > 0 Then
              Call MENSERR(24, "Informe no Editable")
              GoTo 1405
            End If
            TITOT = CVS(Mid$(ZZ$, 161, 4))
            HORAS% = Int(TITOT / 60)
            MINUS% = Int(TITOT - 60 * HORAS%)
            SEGUS% = 60 * (TITOT - 60 * HORAS% - MINUS%)
            While SEGUS% > 59: SEGUS% = SEGUS% - 60: MINUS% = MINUS% + 1: Wend
            While MINUS% > 59: MINUS% = MINUS% - 60: HORAS% = HORAS% + 1: Wend
            '
            XX$ = REGBLAN$("OPERFAIN")
            Call REPLA(XX$, Left$(ZZ$, 64), 1, 64)
            Call REPLA(XX$, MKD$(CAREA#), 65, 8)
            Call REPLA(XX$, MAQUI$, 73, 6)
            Call REPLA(XX$, MKI$(FERE%), 79, 2)
            Call REPLA(XX$, MKI$(HORE%), 81, 2)
            Call REPLA(XX$, MKI$(HORAS%) + Chr$(MINUS%) + Chr$(SEGUS%), 84, 4)
            Call REPLA(XX$, LEGOPER$, 118, 6)
            Call REPLA(XX$, MKS$(REOPER&), 125, 4)
            '
            JJ& = JJ& + 1
            Call GRAREG("!OPERFAIN", XX$, JJ&)
         End If
       End If
     Next KK%
     Call SETULTREG("!OPERFAIN", JJ&)
     '
     If JJ& < 1 Then
        Call MENSERR(24, "No Hay Operaciones Informadas\para la " + NROOF$)
        GoTo 1405
     End If
     '
     TITUPA$ = VALACT2$("ECORDEF")
     PPXX% = InStr(TITUPA$, " - ")
     If PPXX% < 1 Then PPXX% = 17
     CODPIE$ = TRIM$(Left$(TITUPA$, PPXX% - 1))
     CIPIE% = CODINT%(CODPIE$)
     If CIPIE% < 1 Then
        Call MENSERR(24, "Imposible Determinar Pieza de esta O/F")
     End If
     '
   ' MODIFICAR PARA PODER SOLO CAMBIAR LA CANTIDAD Y EL TIEMPO EMPLEADO
   ' AL GRABAR QUE RECALCULE EL TIEMPO STANDARD
     '
     VTB% = VENTABU%("INFOPER/TITUPAN=Informe de Produccion: " + TITUPA$)
     '
     If VTB% >= 0 Then
        '
        While Len(RFF$) < 4 * ULTREG&("!OPERFAIN")
           RFF$ = RFF$ + Chr$(0)
        Wend
        '
        For U& = 1 To ULTREG&("!OPERFAIN")
           '
           XX$ = REGLEIDO$("!OPERFAIN", U&)
           HMS$ = Mid$(XX$, 84, 4)
           TITOT = 60 * CVI(Left$(HMS$, 2)) + Asc(Mid$(HMS$, 3, 1)) + Asc(Mid$(HMS$, 4, 1)) / 60
           Call REPLA(XX$, MKS$(TITOT), 84, 4)
           Call REPLA(XX$, NROOF$, 1, 12)
           Call REPLA(XX$, MKI$(CIPIE%), 13, 2)
           '
           REOPER& = CVS(Mid$(XX$, 125, 4))
           '
           If REOPER& > 0 And REOPER& <= ULTREG&("OPERFAB") Then
               ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
               NORFA$ = Left$(ZZ$, 12)
               If NORFA$ <> NROOF$ Then
                   Call CIERRARCH("*.*")
                   Call MENSERR(24, "Imposible Grabar")
                   Call FINAL("")
               End If
               CAREA# = CVD(Mid$(XX$, 65, 8))
               Call REPLA(ZZ$, MKD$(CAREA#), 87, 8)
               MAQUI$ = Mid$(XX$, 73, 6)
               Call REPLA(ZZ$, MAQUI$, 151, 6)
               FERE% = CVI(Mid$(XX$, 79, 2))
               Call REPLA(ZZ$, MKI$(FERE%), 157, 2)
               HORE% = CVI(Mid$(XX$, 81, 2))
               Call REPLA(ZZ$, MKI$(FERE%), 159, 2)
               LEGOPER$ = Mid$(XX$, 118, 6)
               Call REPLA(ZZ$, LEGOPER$, 165, 6)
               Call REPLA(ZZ$, MKS$(TITOT), 161, 4)
               '
               NOPE% = CVI(Mid$(ZZ$, 15, 2))
               SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREA#
               If SALRE# < 0 Then SALRE# = 0
               NUOROP% = Asc(Mid$(ZZ$, 67, 1))
               TSTAN = TIESTAN(CIPIE%, NOPE%, CAREA#, NUOROP%)
               '
               Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
               Call REPLA(ZZ$, MKS$(TSTAN), 173, 4)
               '
               Call GRAREG("OPERFAB", ZZ$, REOPER&)
           End If
           '
        Next U&
     End If
     '
     Call CIERRARCH("OPERFAB")
     GoTo 1405
     '
End Sub
'
Sub CARGAINF(NROOF$, NOPE%)
     '
     HOII% = HOY(HOS$)
     AGREGA% = 0
     '
     Call ABREORFAB
     '
     sqlx$ = " SELECT * FROM OPERFAB "
     sqlx$ = sqlx$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     sqlx$ = sqlx$ + " AND NumeOper = " & NOPE% & " "
     '
     Set OPFABTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
     With OPFABTMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No se Puede Procesar la Información\Re-Genere la Orden de Fabricación.")
         Exit Sub
       End If
       On Error GoTo 0
       '
       CI% = !Cod_Inte
       DOPE$ = !DescOper
       CAPROG# = !CantProg
       CAREAL# = 0    'CAPROG#    ' !CantReal
       TIETOT = !TiemTota
       TIEPEN = !TiemPend
       MAQUI$ = !Maquina
       FEOP% = CVINT(!TermiPro)
       If FEOP% < 1 Then
           FEOP% = HOII%
           HOOP% = HORANUM(Time$)
         Else
           FEOPX$ = Format$(!TermiPro, "dd/mm/yy  hh:mm:ss")
           FEOP% = FECHANUM(Left$(FEOPX$, 8))
           HOOP% = HORANUM(Mid$(FEOPX$, 11, 8))
       End If
       MALIB$ = "S"
       
     End With
     '
     sqlx$ = " SELECT * FROM INFOFAB "
     sqlx$ = sqlx$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     sqlx$ = sqlx$ + " AND NumeOper = " & NOPE% & " "
     '
     CAINFANT% = 0
     Set INFOFABTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           CAREAL# = CAREAL# + !CantReal
           TIEINF = TIEINF + !TiemEjec
           CAINFANT% = CAINFANT% + 1
         .MoveNext
         Loop
       End If
     End With
     '
     SALDO# = CAPROG# - CAREAL#
     If SALDO# < 0 Then
       SALDO# = 0
       TIEPEN = 0
     End If
     HORAS% = Int(TIEPEN / 60)
     MINUS% = Int(TIEPEN - 60 * HORAS%)
     SEGUS% = 60 * (TIEPEN - 60 * HORAS% - MINUS%)
     While SEGUS% > 59: SEGUS% = SEGUS% - 60: MINUS% = MINUS% + 1: Wend
     While MINUS% > 59: MINUS% = MINUS% - 60: HORAS% = HORAS% + 1: Wend
     OBSE$ = Space$(30)
     '
     YY$ = REGBLAN$("INFOFAB")
     Call REPLA(YY$, NROOF$, 1, 12)
     Call REPLA(YY$, MKI$(CI%), 13, 2)
     Call REPLA(YY$, MKI$(NOPE%), 15, 2)
     Call REPLA(YY$, DOPE$, 17, 48)
     Call REPLA(YY$, MKD$(SALDO#), 65, 8)
     Call REPLA(YY$, MAQUI$, 73, 6)
     Call REPLA(YY$, MKI$(FEOP%), 79, 2)
     Call REPLA(YY$, MKI$(HOOP%), 81, 2)
     Call REPLA(YY$, MALIB$, 83, 1)
     Call REPLA(YY$, MKI$(HORAS%) + Chr$(MINUS%) + Chr$(SEGUS%), 84, 4)
     Call REPLA(YY$, OBSE$, 88, 30)
     Call REPLA(YY$, String$(8, 0), 118, 8)
     '
1420 XX$ = OBJPLA$("INFOFAB", YY$)
     If XX$ = "" Then
       NOPE% = 0
       Exit Sub
     End If
     '
     YY$ = XX$
     CAINF# = CVD(Mid$(XX$, 65, 8))
     '
     MAQUI$ = Mid$(XX$, 73, 6)
     If TRIM$(ECOARCH$("PADMAQ", MAQUI$)) = "" Then
         Call MENSERR(24, " Equipo Incorrecto.")
         GoTo 1420
     End If
     '
     FEOP% = CVI(Mid$(XX$, 79, 2))
     HOOP% = CVI(Mid$(XX$, 81, 2))
     If FEOP% > HOII% Then
           Call MENSERR(24, "Fecha Incorrecta.")
           GoTo 1420
         ElseIf FEOP% = HOII% Then
           If HOOP% > HORANUM(Time$) Then
             Call MENSERR(24, "Hora Incorrecta")
             GoTo 1420
           End If
     End If
     FEX = FEOP%
     FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEX$(HOOP%))
     '
     If ULTREG&("PADRON") > 0 Then
        If CVI(Mid$(XX$, 118, 2)) < 1 Then
           Call MENSERR(24, "Falta Informar Operario(s)")
           GoTo 1420
        End If
        For KK% = 118 To 125 Step 2
           LEGA% = CVI(Mid$(XX$, KK%, 2))
           If LEGA% > 0 Then
              If ECOARCH$("PADRON", MKI$(LEGA%)) = "" Then
                 Call MENSERR(24, "Legajo Operario no Válido")
                 GoTo 1420
              End If
           End If
        Next KK%
     End If
     '
     MALIB$ = " "
     If CAINF# >= SALDO# - 0.05 Then
         MALIB$ = "S"
     End If
     Call REPLA(XX$, MALIB$, 83, 1)
     OBSE$ = TRIM$(Mid$(XX$, 88, 30))
     '
     TIEINF = 60 * CVI(Mid$(XX$, 84, 2)) + Asc(Mid$(XX$, 86, 1)) + Asc(Mid$(XX$, 87, 1)) / 60
     '////***NUEVO***/////
     If TIEINF <= 0 Then 'PARA evitar OVERFLOW
        Call COMUNI("Debe Ingresar un Tiempo Expresado en: \ Horas, Minutos y Segundos")
        GoTo 1420
     End If
     Call REPLA(XX$, MKS$(TIEINF), 84, 4)
     TIEINF = CVS(Mid$(XX$, 84, 4))
     '
     LEGOPER$ = Mid$(XX$, 118, 8)
     '
     STATUOP% = 3 ' Cumplida
     If CAINF# < SALDO# - 0.05 Then
       STATUOP% = 1   ' AUN PENDIENTE
       TTXX$ = "Queda un Saldo Pendiente\para esta Operación\\Mantener Pendiente\Cancelar Saldo"
       If COMALTER%(TTXX$) = 2 Then
         STATUOP% = 8 ' CANCELADO
       End If
     End If
     '
     Call ABREORFAB
     With InfoFab
       .AddNew
        !IdSubOr = NROOF$
        !DescriOF = ECOARCH$("ECORDEF", NROOF$)
        !NuOrInf = CAINFANT%
        !Cod_Inte = CI%
        !Cod_Exte = CODEXT$(CI%)
        !Descrip = DESCRIT$(CI%)
        !NumeOper = NOPE%
        !AltOper = 0
        !ReTraba = 0
        !DescOper = DOPE$
        !CanOrPro = CAPROG#
        !CantReal = CAINF#
        !Maquina = MAQUI$
        !DescMaq = ECOARCH$("PADMAQ", MAQUI$)
        !LegaOps = LEGOPER$
        !FeHorLib = FEHOLI
        !Hora_Des = 0
        !Hora_Has = HOOP%
        !MaqLiber = MALIB$
        !TiemEjec = TIEINF
           MODO% = CAINFANT%
        !TiemStan = TIESTAN(CI%, NOPE%, CAINF#, MODO%)   ' 1 o mas es para que no considere preparacion
           'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
           ' A LOS CAMPOS DEL REGISTRO. "!"
           'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
           RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
           If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = OBSE$
        !Referen = " "
       .Update
     End With
     InfoFab.Close
     '
     sqlx$ = " SELECT * FROM INFOFAB "
     sqlx$ = sqlx$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     sqlx$ = sqlx$ + " AND NumeOper = " & NOPE% & " "
     '
     CAREAL# = 0: TIEINF = 0
     Set INFOFABTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           CAREAL# = CAREAL# + !CantReal
           TIEINF = TIEINF + !TiemEjec
         .MoveNext
         Loop
       End If
     End With
     '
     With OperFab
       .FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
       If .NoMatch = False Then
         .Edit
         !CantReal = CAREAL#
         !CantPend = !CantProg - CAREAL#
           If !CantPend < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CantPend = 0
           CANTPE# = !CantPend
           MODO% = CAINFANT%
         !TiemEjec = TIEINF
         !TiemPend = TIEMAQUI(CI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
           If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
         !StatOper = STATUOP%
         .Update
       End If
     End With
     '
End Sub
'

Sub GESALPROP(MODO%)
   '
   Dim NORFAX$(4096)
   NROOFA$ = "": NOPEA% = 0
   '
   FIN& = ULTREG&("OPERFAP")
   W& = 0
   For U& = 1 To FIN&
      XX$ = REGLEIDO$("OPERFAP", U&)
      CAPROD# = CVD(Mid$(XX$, 79, 8))
      CAREA# = CVD(Mid$(XX$, 87, 8))
      SALDO# = CAPROD# - CAREA#
      If SALDO# >= 0.05 Then
         NROOF$ = Left$(XX$, 12)
         CI% = CVI(Mid$(XX$, 13, 2))
         NOPE% = CVI(Mid$(XX$, 15, 2))
         DOPE$ = Mid$(XX$, 17, 48)
         '
         For KX& = 1 To W&
            If NORFAX$(KX&) = NROOF$ Then GoTo 10
         Next KX&
         '
         z$ = REGBLAN$("SPPROCE")
         Call REPLA(z$, MKI$(CI%), 1, 2)
         Call REPLA(z$, NROOF$, 3, 12)
         Call REPLA(z$, MKI$(CI%), 15, 2)
         Call REPLA(z$, MKI$(NOPE%), 17, 2)
         Call REPLA(z$, DOPE$, 19, 48)
         Call REPLA(z$, MKD$(CAPROD#), 67, 8)
         W& = W& + 1
         Call GRAREG("SPPROCE", z$, W&)
         NORFAX$(W&) = NROOF$: NOPEA% = NOPE%
         GoTo 19
         '
10       z$ = REGLEIDO$("SPPROCE", KX&)
         NOPEA% = CVI(Mid$(z$, 17, 2))
         If NOPE% < NOPEA% Then
            Call REPLA(z$, MKI$(NOPE%), 17, 2)
            Call REPLA(z$, DOPE$, 19, 48)
            Call GRAREG("SPPROCE", z$, KX&)
         End If
      End If
19 Next U&
   '
   Call SETULTREG("SPPROCE", W&)
   Call CIERRARCH("SPPROCE")
   '
   If MODO% = 1 Then
         Call FILESORT("SPPROCE", "", 3, 4, "ASC")
         Call CONECRUN("*LISPPR1", "Saldos en Proceso por O.F")
         'Call CONECRUN("*SOSPPR1", "Saldos en Proceso por O.F")
       ElseIf MODO% = 2 Then
         Call FILESORT("SPPROCE", "", 1, 3, "ASC")
         Call CONECRUN("*LISPPR2", "Saldos en Proceso por Producto")
         'Call CONECRUN("*SOSPPR2", "Saldos en Proceso por Producto")
   End If
   '
End Sub

Private Sub InfMayRa_Click()
Call Command15_Click
End Sub

Private Sub MaArtStck_Click()
Call Command9_Click
End Sub

Private Sub MaEq_Click()
Call Command5_Click
End Sub

Private Sub OpPenOF_Click()
Call Command7_Click
End Sub

Private Sub OrdFabri_Click()
Call Command4_Click
End Sub

Private Sub OrdFabric_Click()
Call Command2_Click
End Sub

Private Sub OtAccDirec_Click()
   Call ACCDIR("INFOFAB")
End Sub

Private Sub PadOper_Click()
Call Command19_Click
End Sub

Private Sub ProgDia_Click()
Call Command6_Click
End Sub

Private Sub Progmaqui_Click()
Call Command25_Click
End Sub

Private Sub Salir_Click()
Call Command1_Click
End Sub

Private Sub SecRuOp_Click()
Call Command9_Click
End Sub

Private Sub TiOpCum_Click()
Call Command30_Click
End Sub
'
Sub GENRESDAT()
    '
    Call ABRESTOCKS
    '
    sqlx$ = " SELECT * FROM RESERVA "
    sqlx$ = sqlx$ + " WHERE CantRes-CantCons > 0 "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    '
    Set RESERTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      'If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
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
    On Error Resume Next
    RESERTMP.Close
    Set RESERTMP = Nothing
    On Error GoTo 0
    '
    Call ABREORFAB
    '
    sqlx$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR "
    sqlx$ = sqlx$ + " WHERE StatuOrf < 3 "
    sqlx$ = sqlx$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.")   '  Imposible Continuar la Carga.")
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
    On Error Resume Next
    ORDEFABTMP.Close
    Set ORDEFABTMP = Nothing
    On Error GoTo 0
End Sub
'\\\OT-06-0252-WAR-26/06/06///
Sub INFAPROY()      ' INFORME DE PROYECTOS - GENERAL
     '
100  COPROY$ = ""
     '
     Call GENINDIPROY
     '
     If ULTREG&("CATOPER") <= 0 Then
       Call MENSERR(24, "No se Puede Valorizar el Informe\No Existen Definidas Categorias de Empleados")
       GoTo 99
     End If
     '
1405 Call FRESHECHO("")
     If ULTREG&("!SELPROY") < 1 Then
         Call MENSERR(24, "No hay Proyectos en Curso.")
         GoTo 99
     End If
     '
     Screen.MousePointer = 11
     '2.-SACO NUMERO DE PROXIMO INFORME
     Call ABREORFAB
     sqlx$ = ""
     sqlx$ = "SELECT * FROM INFOPROY "
     sqlx$ = sqlx$ + " Order by Nuor_Inf ASC"
     Set INFTMP = Ofabric.OpenRecordset(sqlx$, 4)
     With INFTMP
      On Error Resume Next
      .MoveLast
      If Err Then
        PRONU1% = 1
        PRONU2$ = "IF." + TRIM$(Str$(PRONU1%))
        On Error GoTo 0
        GoTo 1412
      End If
      On Error GoTo 0
      PRONU1% = !Nuor_Inf + 1
      PRONU2$ = "IF." + TRIM$(Str$(PRONU1%))
     End With
     '
1412 HOII% = HOY(HOSS$)
     FEOP% = HOII%
     HOOP% = HORANUM(Time$)
     '
     YY$ = Space$(128)
     Call REPLA(YY$, PRONU2$, 1, 12)
     Call REPLA(YY$, COPROY$, 13, 8)
     Call REPLA(YY$, MKI$(FEOP%), 70, 2)
     Call REPLA(YY$, MKI$(HOOP%), 72, 2)
     Call REPLA(YY$, String$(8, 0), 62, 8)
     Call REPLA(YY$, String$(8, 0), 74, 2)
     Call REPLA(YY$, String$(8, 0), 76, 1)
     Call REPLA(YY$, String$(8, 0), 77, 1)
     '
1413 XX$ = OBJPLA$("INFOPROY", YY$)
     If XX$ = "" Then
       Screen.MousePointer = 1
       GoTo 99
     End If
     YY$ = XX$
     '
     COPROY$ = TRIM$(Mid$(XX$, 13, 8))
     If ECOARCH$("INDIPROY", COPROY$) = "" Then
        Call MENSERR(24, "Proyecto no Válido")
        GoTo 1413
     End If
     '
     FF12% = CVI(Mid$(XX$, 70, 2))
     If VALIPROY%(COPROY$, FF12%) = (-1) Then
        Call MENSERR(24, "Proyecto no Vigente \para la Fecha Seleccionada")
        GoTo 1413
     End If
     '
     REFE1$ = TRIM$(Mid$(XX$, 21, 41))
     If REFE1$ = "" Then
        Call MENSERR(24, "Referencia no Válida")
        GoTo 1413
     End If
     '
     If ULTREG&("PADRON") > 0 Then
        If CVI(Mid$(XX$, 62, 2)) < 1 Then
           Call MENSERR(24, "Falta Informar Operario(s)")
           GoTo 1413
        End If
        For KK% = 62 To 68 Step 2
           LEGA1% = CVI(Mid$(XX$, KK%, 2))
           If LEGA1% > 0 Then
              If ECOARCH$("PADRON", MKI$(LEGA1%)) = "" Then
                 Call MENSERR(24, "Legajo Operario no Válido")
                 GoTo 1413
              End If
           End If
        Next KK%
     End If
     '
     TIEINF = 60 * CVI(Mid$(XX$, 74, 2)) + Asc(Mid$(XX$, 76, 1)) + Asc(Mid$(XX$, 77, 1)) / 60
     If TIEINF <= 0 Then 'PARA evitar OVERFLOW
        Call COMUNI("Debe Ingresar un Tiempo Expresado en: \ Horas, Minutos y Segundos")
        GoTo 1413
     End If
     Call REPLA(XX$, MKS$(TIEINF), 74, 4)
     TIEINF = CVS(Mid$(XX$, 74, 4))
     '
     LEGOPER$ = Mid$(XX$, 62, 8)
     '
     Dim LEGA%(512)
     Dim NOME$(512)
     Dim CATE%(512)
     '
     For U& = 1 To ULTREG&("PADRON")
       XX$ = REGLEIDO$("PADRON", U&)
       LEGA%(U&) = CVI(Left$(XX$, 2))
       NOME$(U&) = Mid$(XX$, 3, 30)
       CATE%(U&) = Val(Mid$(XX$, 54, 1))
       CAOPE% = U&
       If U& >= 512 Then GoTo 20
     Next U&
     '
20   For KK% = 1 To Len(LEGOPER$) Step 2
      LEGX% = CVI(Mid$(LEGOPER$, KK%, 2))
      If LEGX% > 0 Then
        For MM% = 1 To CAOPE%
          If LEGA%(MM%) = LEGX% Then
            CATOPX% = CATE%(MM%)
            VALOMOD = VALHORA(CATOPX%) * (TIEINF * 60) / 3600
            TOTAVAL1 = VALOMOD + TOTAVAL1
            GoTo 39
          End If
        Next MM%
      End If
39   Next KK%
     '
     Call ABREORFAB
     With InfoProy
       .AddNew
         !Nuor_Inf = PRONU1%
         !Nume_Inf = PRONU2$
         !Fecha_Inf = CVDAT(FF12%)
         !Refer_Inf = REFE1$
         !Cant_Inf = TIEINF 'MINUTOS
         !Proy_Inf = COPROY$
         !LegaOps_Inf = LEGOPER$
         !Total_Inf = TOTAVAL1
       .Update
     End With
     InfoProy.Close
     '
     Call COMUNI("Informe Procesado")
     GoTo 100
99   Screen.MousePointer = 1
     '
End Sub
Sub GENINDIPROY()
  '
  REBLA1$ = REGBLAN$("INDIPROY")
  JJ& = 0
  HOII% = HOY(HO$)
  Call COPYSTRU("INDIPROY", "SELPROY")
  sqlx$ = "SELECT * FROM TAPROYEC"
  sqlx$ = sqlx$ + " WHERE STAT_PROY < 3 "
  Call ABREPLANCUEN
  Set CIETMP = Contable.OpenRecordset(sqlx$, 4)
  With CIETMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    Do While .EOF = False
     XZ$ = REBLA1$
     Call REPLA(ZX$, !Cod_Proy, 1, 8)
     Call REPLA(ZX$, TRIM$(ABRECLI$(!Clien_Proy)) & " - " & !Deno_Proy, 9, 48)
     JJ& = JJ& + 1
     Call GRAREG("!SELPROY", ZX$, JJ&)
     .MoveNext
    Loop
  End With
99 Call SETULTREG("!SELPROY", JJ&)
  '
End Sub
Sub ANUINFAPROY()
  '
  Screen.MousePointer = 11
  Call FRESHECHO("INFAPROY")
  '
  If DERACCE%("ANUINFO", "Anulacion de Informes de Proyecto") < 2 Then
    GoTo 99
  End If
  '
  Call ABREORFAB
  sqlx$ = ""
  sqlx$ = "SELECT * FROM INFOPROY "
  sqlx$ = sqlx$ + " Order by Nuor_Inf ASC"
  Set INFTMP = Ofabric.OpenRecordset(sqlx$, 4)
  With INFTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call MENSERR(24, "No Existen Informes\para su Anulación")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    JJ& = 0
    REBLA$ = REGBLAN$("INFAPROY")
    Do While .EOF = False
      YY$ = REBLA$
      Call REPLA(YY$, MKS$(!Nuor_Inf), 1, 4)
      REFER1$ = ""
      F1% = CVINT(!Fecha_Inf)
      F2$ = FECHATEX$(F1%)
      REFER1$ = F2$ + " - " + !Refer_Inf
      Call REPLA(YY$, REFER1$, 5, 60)
      JJ& = JJ& + 1
      Call GRAREG("INFAPROY", YY$, JJ&)
      .MoveNext
    Loop
    Call SETULTREG("INFAPROY", JJ&)
  End With
  '
5 Call SELECHO("INFAPROY")
  NUINF1& = XVALO(VALACT1$("INFAPROY"))
  If NUINF1& > 0 Then
    OPX1% = COMALTER%("¿ Confirma Anular el Informe: " + "IF." + TRIM$(Str$(NUINF1&)) + " ?\\No, Cancelar\Si, Cerrar")
    If OPX1% = 2 Then
      With InfoProy
        .FindFirst "Nuor_Inf = " & NUINF1& & ""
        If .NoMatch = False Then
          .Delete
        End If
      End With
      Call COMUNI("Anulación Completa")
    End If
  End If
  '
99 Screen.MousePointer = 1
End Sub
'\\\OT-06-0252-WAR-FIN///

Sub CALCUPROCE0()
     'PARA DESARROLLAR SALDOS DE PRODUCTOS EN PROCESO - 22/08/06
     '\\\OT-06-0335-22/08/06///
     Dim NOFA$() ' vector donde cargo ordenes fabricacion
     '
     Call ABREORFAB ' funcion que abre bases de datos
     sqlx$ = " SELECT * FROM ordefabr "
     sqlx$ = sqlx$ + " WHERE  statuorf < 3 "
     Set PROXFECTMP = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
     X1% = 0 ' para establecer posicion en el vector(NOFA$)
     '
     With PROXFECTMP
        On Error Resume Next 'si hubiera error salta a la sig linea
        .MoveFirst
        If Err Then ' en caso de error...
          On Error GoTo 0  ' cancela Resume Next
          Call MENSERR(24, "No hay ordenes de Fabricación en Proceso")
          GoTo 10          ' salta a la linea 10
        End If
        On Error GoTo 0    'Cancela el Resume Next
        '
        Do While Not .EOF  'mientras no sea el final
          X1% = X1% + 1    ' incremento pos del vector
          ReDim Preserve NOFA$(X1%)  'redimensiono vector
          'con preserve redimensiono "preservando" los datos
          NOFA$(X1%) = TRIM$(!IdSubOr)
          'en la posicion X1% escribo "idsubor"
        .MoveNext ' avanzo al siguiente
        Loop
     End With
     '
10   JJ& = 0 ' Tiene nº de registro en acuproce
     REBLA$ = REGBLAN$("ACUPROCE") '
     'crea un string de blancos como la long del archivo
     With InfoFab
        '
        On Error Resume Next
         .MoveFirst '
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No hay Informes de Fabricación")
          GoTo 70
        End If
        On Error GoTo 0
        '
        Do While .EOF = False
          '
          XX$ = REBLA$
          OBSEX$ = "Re-Trabajo"
          If IsNull(!ReTraba) Or !ReTraba < 1 Then OBSEX$ = ""
          If TRIM$(!Maquina) = "Ajuste" Then OBSEX$ = "Ajuste"
          '
          NROOF$ = "" ' NRO ORDEN FABRICACION
          If IsNull(!IdSubOr) = False Then
             NROOF$ = TRIM$(!IdSubOr)
             'si no esta vacio idsubor sera el valor de NROOF$
          End If
          '
          If NROOF$ <> "" Then
             For KI% = 1 To X1% 'hasta total vector
               If NROOF$ = NOFA$(KI%) Then '
                 GoTo 20 ' ENCONTRÓ LA of PENDIENTE
               End If
             Next KI%
             GoTo 60
          End If
          '
20        ALTOP% = 0
          On Error Resume Next
          ALTOP% = !AltOper
          On Error GoTo 0
          '
  'repla va escribiendo los datos
          Call REPLA(XX$, MKI$(!Cod_Inte), 1, 2)
          Call REPLA(XX$, !IdSubOr, 3, 12)
          Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
          Call REPLA(XX$, DOPE$, 17, 44)
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 61, 2)
            HDES% = 0: HHAS% = 0
            On Error Resume Next
            HDES% = !Hora_Des
            HHAS% = !Hora_Has
            On Error GoTo 0
          Call REPLA(XX$, MKI$(HDES%), 63, 2)
          Call REPLA(XX$, MKI$(HHAS%), 65, 2)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 67, 4)
          Call REPLA(XX$, MKD$(!CantReal), 129, 8)
          Call REPLA(XX$, OBSEX$, 169, 32)
          '
          JJ& = JJ& + 1 ' incrementa registro ACUPROCE
          Call GRAREG("ACUPROCE", XX$, JJ&)
          'Graba en acuproce, lo que hay en xx en pos jj
60      .MoveNext
        Loop
     End With
     '
70   ' AGREGA ENTRADAS A STOCK POR CIERRE DE o/fABRIC
     Call ABRESTOCKS
     sqlx$ = " SELECT * FROM Movisto "
     sqlx$ = sqlx$ + " WHERE  CodiMovi = 'OF' "
     Set Movisel = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
     With Movisel
        '
        On Error Resume Next
         .MoveFirst '
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Movimientos de Cierre de O/F")
          GoTo 90
        End If
        On Error GoTo 0
        '
        Do While .EOF = False
          XX$ = REBLA$
          NROOF$ = TRIM$(!DOPRILAR)
          If NROOF$ <> "" Then
             For KI% = 1 To X1% 'hasta total vector
               If NROOF$ = NOFA$(KI%) Then '
                 GoTo 75 ' ENCONTRÓ LA of PENDIENTE
               End If
             Next KI%
             GoTo 80
          End If
          '
75        Call REPLA(XX$, MKI$(!Cod_Inte), 1, 2)
          Call REPLA(XX$, NROOF$, 3, 12)
          Call REPLA(XX$, MKI$(9999), 15, 2)
          Call REPLA(XX$, "9999   Parcial O/F - Ingreso Stock", 17, 44)
          Call REPLA(XX$, MKI$(CVINT(!FechMov)), 61, 2)
          Call REPLA(XX$, MKI$(0), 63, 2)
          Call REPLA(XX$, MKI$(0), 65, 2)
          Call REPLA(XX$, MKS$(0), 67, 4)
          Call REPLA(XX$, MKD$(!CantIngre), 129, 8)
          '
          JJ& = JJ& + 1 ' incrementa registro ACUPROCE
          Call GRAREG("ACUPROCE", XX$, JJ&)
          'Graba en acuproce, lo que hay en xx en pos jj
80      .MoveNext
        Loop
     End With
     '
     
90   Call SETULTREG("ACUPROCE", JJ&) 'Pone marca de fin archivo
     'JJ& es el ultimo registro valido
     Call CIERRARCH("ACUPROCE")
     'cierra el archivo
     '
     'ordena
     Call HEADERS("ACUPROCE", "")
     Call FILESORT("ACUPROCE", "", 4, 4, "ASC")
     Call FILESORT("ACUPROCE", "", 3, 3, "ASC")
     Call FILESORT("ACUPROCE", "", 1, 1, "ASC")
     '
End Sub

Sub CALCUPROCE1()
     '
     Screen.MousePointer = 11
     '
     CI1% = 0     ' CODIGOS
     NROF1$ = ""  ' NUMERO DE O/F
     NOP1% = 0    ' NUMERO DE OPERACION
     '
     JJ1& = 0 ' CANT REG ACUPROC1
     '
     FIN& = ULTREG&("ACUPROCE")
     For II& = 1 To FIN&
        Call TRACE(20, II&, FIN&)
        X$ = REGLEIDO$("ACUPROCE", II&)
        If TRIM$(Mid$(X$, 169, 32)) <> "Re-Trabajo" Then
            'Me cargo el registro en esta variable
           CI% = CVI(Mid$(X$, 1, 2))          ' PARA CODIGO
           NROF$ = TRIM$(Mid$(X$, 3, 12))     ' PARA Nº Orden Fabric
           NOP% = CVI(Mid$(X$, 15, 2))        ' Para Nº Orden de Pedido
           If NOP% > 0 Then
              CANTU# = CVD(Mid$(X$, 129, 8))
              HORAS = CVS(Mid$(X$, 67, 4))
              DESCOP$ = DESCRIOP$(CI%, NOP%, 0)
              If NOP% = 9999 Then DESCOP$ = "Parcial O/F - Ingreso Stock"
              '
              If CI% = CI1% Then
                If NROF$ = NROF1$ Then
                  If NOP% = NOP1% Then
                    YY$ = REGLEIDO$("ACUPROC1", JJ1&)
                    TotCant# = CVD(Mid$(YY$, 129, 8)) + CANTU#
                    TOTHORAS = CVS(Mid$(YY$, 67, 4)) + HORAS
                    Call REPLA(YY$, MKD$(TotCant#), 129, 8)
                    Call REPLA(YY$, MKS$(TOTHORAS), 67, 4)
                    Call GRAREG("ACUPROC1", YY$, JJ1&)
                    GoTo 79
                  End If
                End If
              End If
              '
              XX$ = REGBLAN$("ACUPROC1")
              Call REPLA(XX$, MKI$(CI%), 1, 2)
              Call REPLA(XX$, NROF, 3, 12)
              Call REPLA(XX$, MKI$(NOP%), 15, 2)
              Call REPLA(XX$, DESCOP$, 17, 44)
              Call REPLA(XX$, MKS$(HORAS), 67, 4)
              Call REPLA(XX$, MKD$(CANTU#), 129, 8)
              JJ1& = JJ1& + 1 ' incrementa registro
              Call GRAREG("ACUPROC1", XX$, JJ1&)
              '
              CI1% = CI%
              NROF1$ = NROF$
              NOP1% = NOP%
           End If
           'vuelvo a asignar igualdades para comparar despues
        End If
79   Next II&
     '
     Call SETULTREG("ACUPROC1", JJ1&) 'Pone marca de fin archivo
          'JJ& es el ultimo registro valido
     Call CIERRARCH("ACUPROC1")
          'cierra el archivo
     Screen.MousePointer = 1

End Sub

Sub CALCUPROCE2()
     '
     Screen.MousePointer = 11
     '
     CI1% = 0     ' CODIGOS
     NROFA$ = ""  ' NUMERO DE O/F
     NOP1% = 0    ' NUMERO DE OPERACION
     '
     JJ1& = 0 ' CANT REG ACUPROC1
     '
     FIN& = ULTREG&("ACUPROC1")
     For II& = 1 To FIN&
        Call TRACE(20, II&, FIN&)
        XX$ = REGLEIDO$("ACUPROC1", II&)
         'Me cargo el registro en esta variable
        CI% = CVI(Mid$(XX$, 1, 2))          ' PARA CODIGO
        NROF$ = TRIM$(Mid$(XX$, 3, 12))     ' PARA Nº Orden Fabric
        If NROF$ <> "" Then                ' EXCLUYE LOS SIN NRO DE o/f
          '
          CANTU# = CVD(Mid$(XX$, 129, 8))
          NOP% = CVI(Mid$(XX$, 15, 2))        ' Para Nº DE OPERACION
          '
          If NROF$ = NROFA$ Then
             YY$ = REGLEIDO$("ACUPROC2", JJ1&)
             CANPE# = CVD(Mid$(YY$, 129, 8)) - CANTU#
             Call REPLA(YY$, MKD$(CANPE#), 129, 8)
             Call REPLA(YY$, "", 145, 32)
             If CANPE# <= (-0.05) Then
               ' PONER OBSERVACION = "NEGATIVO"
               Call REPLA(YY$, "Negativo", 145, 32)
             End If
             Call GRAREG("ACUPROC2", YY$, JJ1&)
          End If
          '
          ' AGREGA UN REGISTRO
          If NOP% <> 9999 Then   ' PARA EXCLUIR ENTRADAS A STOCK
            JJ1& = JJ1& + 1
            Call GRAREG("ACUPROC2", XX$, JJ1&)
          End If
          '
          NROFA$ = NROF$
        End If
        '
79   Next II&
     '
     Call SETULTREG("ACUPROC2", JJ1&) 'Pone marca de fin archivo
     Call CIERRARCH("ACUPROC2")
     '
     Screen.MousePointer = 1
     '
End Sub

Sub DEFCELULAS()
    '
    Call TRALOCAL("ICELULAS", "")
    VTB% = VENTABU%("ICELULAS/NC/TITUPAN=Indice de Células Productivas")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    TTXX$ = ""
    For i& = 1 To ULTREG&("!ICELULAS")
       X$ = REGLEIDO$("!ICELULAS", i&)
       If CVI(Left$(X$, 2)) > 0 Then
         IGRUPO% = CVI(Left$(X$, 2))
         For JJ& = 1 To Len(TTXX$) Step 2
           If CVI(Mid$(TTXX$, JJ&, 2)) = IGRUPO% Then GoTo 19
         Next JJ&
         J& = J& + 1
         Call GRAREG("ICELULAS", X$, J&)
         TTXX$ = TTXX$ + MKI$(IGRUPO%)
       End If
19  Next i&
    Call SETULTREG("ICELULAS", J&)
    Call CIERRARCH("ICELULAS")
    '
End Sub

Sub CARCELULAS()
    '
10  Call SELECHO("ICELULAS")
    ICELU% = XVALO(VALACT1$("ICELULAS"))
    If ICELU% < 1 Then GoTo 99
    '
    DCELU$ = ECOARCH$("ICELULAS", MKI$(ICELU%))
    Call FILTERFILE("DCELULAS", "", 1, MKI$(ICELU%))
    VTB% = VENTABU%("DCELULAS/NC/TITUPAN=Célula: " + DCELU$)
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    TTXX1$ = "": TTXX2$ = ""
    For i& = 1 To ULTREG&("!DCELULAS")
       X$ = REGLEIDO$("!DCELULAS", i&)
       If CVI(Mid$(X$, 3, 2)) > 0 Then
         NLEGA% = CVI(Mid$(X$, 3, 2))
         For JJ& = 1 To Len(TTXX1$) Step 2
           If CVI(Mid$(TTXX1$, JJ&, 2)) = NLEGA% Then GoTo 19
         Next JJ&
         Call REPLA(X$, MKI$(ICELU%), 1, 2)
         J& = J& + 1
         Call GRAREG("!DCELULAS", X$, J&)
         TTXX1$ = TTXX1$ + MKI$(NLEGA%)
         If TTXX2$ <> "" Then TTXX2$ = TTXX2$ + " / "
         TTXX2$ = TTXX2$ + TRIM$(Str$(NLEGA%))
       End If
19  Next i&
    Call SETULTREG("!DCELULAS", J&)
    Call CIERRARCH("!DCELULAS")
    '
    Call FILTERUPDATE("DCELULAS", "!DCELULAS", 1, MKI$(ICELU%))
    '
    X$ = FILTERGETRECORD$("ICELULAS", 1, MKI$(ICELU%))
    If X$ <> "" Then
      If X$ <> REGBLAN$("ICELULAS") Then
        Call REPLA(X$, TTXX1$ + String$(64, 0), 33, 64)
        Call REPLA(X$, TTXX2$ + Space$(160), 97, 160)
        RFF$ = RECFILT$("ICELULAS", 1, MKI$(ICELU%))
        RECNO& = CVS(RFF$)
        Call GRAREG("ICELULAS", X$, RECNO&)
      End If
    End If
    Call CIERRARCH("ICELULAS")
    '
    GoTo 10
    '
99  Exit Sub
End Sub
