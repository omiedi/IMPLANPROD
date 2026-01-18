VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INFOFA07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Producción"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9900
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   68
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command39 
      Caption         =   "RECU-GABAL"
      Height          =   315
      Left            =   2880
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "MIGRA INFO PROD"
      Height          =   255
      Left            =   5520
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command34 
      Caption         =   "MIGRA - OPER"
      Height          =   255
      Left            =   7440
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00808080&
      Height          =   6315
      Left            =   9030
      ScaleHeight     =   6255
      ScaleWidth      =   1110
      TabIndex        =   52
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   570
         TabIndex        =   70
         Top             =   4800
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command37 
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   1980
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3240
         Width           =   650
      End
      Begin VB.CommandButton Command36 
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
         Picture         =   "INFOFA07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Informes de Produccion por Rango de Fechas"
         Top             =   2595
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "INFOFA07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Acceso Directo a Maestro de Items de Stock"
         Top             =   1335
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
         Picture         =   "INFOFA07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
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
         Height          =   630
         Left            =   105
         Picture         =   "INFOFA07.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   705
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
         Picture         =   "INFOFA07.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Control de Integridad de la Base de Datos de Operaciones Pendientes"
         Top             =   3960
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MANO DE OBRA"
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
      Left            =   210
      TabIndex        =   42
      Top             =   3780
      Width           =   4110
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   57
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Acceso a Productividad General e Individual por Operario"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Eficiencia de Mano de Obra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   59
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
      Caption         =   "COMPLEMENTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2715
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":17C8
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   525
         ScaleHeight     =   615
         ScaleWidth      =   2610
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":1AD2
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Caption         =   "PRODUCTOS EN PROCESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   350
            ScaleHeight     =   615
            ScaleWidth      =   3300
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
               Height          =   645
               Left            =   2720
               Picture         =   "INFOFA07.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   60
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
               Height          =   645
               Left            =   2130
               Picture         =   "INFOFA07.frx":1F26
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
               Height          =   645
               Left            =   0
               Picture         =   "INFOFA07.frx":2230
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Control de Operaciones Pendientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Control  Productos en Proceso"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   41
               Top             =   105
               Width           =   1455
            End
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1785
         ScaleHeight     =   615
         ScaleWidth      =   1980
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
            Height          =   645
            Left            =   -20
            Picture         =   "INFOFA07.frx":253A
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   615
         ScaleWidth      =   1980
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":2844
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   945
         ScaleHeight     =   615
         ScaleWidth      =   1980
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":2B4E
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   525
         ScaleHeight     =   615
         ScaleWidth      =   1980
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":2E58
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
         Caption         =   "OPERACIONES PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   3135
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
               Height          =   645
               Left            =   0
               Picture         =   "INFOFA07.frx":2FA2
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
               Height          =   645
               Left            =   2560
               Picture         =   "INFOFA07.frx":32AC
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
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
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
      Begin VB.Image Image1 
         Height          =   375
         Left            =   3600
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PARTES DE FABRICACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   3585
         TabIndex        =   29
         Top             =   525
         Width           =   3615
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
            Height          =   645
            Left            =   3045
            Picture         =   "INFOFA07.frx":35B6
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Correccion de Informes de Tiempos de Operaciones Previamente Cargados"
            Top             =   0
            Width           =   590
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
            Height          =   645
            Left            =   2440
            Picture         =   "INFOFA07.frx":3700
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
            Height          =   645
            Left            =   0
            Picture         =   "INFOFA07.frx":3A0A
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
            Picture         =   "INFOFA07.frx":3D14
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFOFA07.frx":3E5E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFOFA07.frx":4092
      TabIndex        =   69
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11025
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu PaFab 
      Caption         =   "Transacciones"
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
      Begin VB.Menu mnuInformePorMotivo 
         Caption         =   "Informes por Motivo"
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
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
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
End
Attribute VB_Name = "INFOFA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AUTOCONTROL()
    '
    '27/03/07 (OT-07-0087)
    '
    FORIPRE$ = TRIM$(Control$("IFABRICA", "FAUTCONT"))
    '
    Call DESHASFECHA(DES1%, HAS1%, PERIO$)
    VDEVU$ = PERIO$
    If VDEVU$ <> "" Then
       FEC1 = FETEXCOR1$(DES1%)
       FEC2 = FETEXCOR1$(HAS1%, "HASTA") ' CON HASTA PARA QUE CONCATENE HORAS
    End If
    '
    Call SELECHO("SECPROD")
    VDEVU1$ = VALACT1$("SECPROD")
    VDEUV2$ = VALACT2$("SECPROD")
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab "
    SQLX$ = SQLX$ + " WHERE  FeHorLib  >=  " & FEC1 & " "
    SQLX$ = SQLX$ + " AND  FeHorLib  <=  " & FEC2 & " "
    SQLX$ = SQLX$ + " ORDER BY INT(FeHorLib) ASC, Hora_Des ASC "
    On Error Resume Next
    'Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PROXFECTMP As ADODB.Recordset
    Set PROXFECTMP = New ADODB.Recordset
    PROXFECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
            EQUI$ = !MAQUINA
          '
            Sec = SECPERTE(EQUI$)
          '
            If Sec = XVALO(Mid(VDEVU1$, 1, 5)) Then
           '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = FECHATEX$(CVINT(!FeHorLib))
              TETABU1$(2, CAITAB1%) = !idsubor
              TETABU1$(3, CAITAB1%) = CODEXT$(!Cod_Inte)
              DOPE$ = TRIM$(Str$(!NUMEOPER))
              TETABU1$(4, CAITAB1%) = DOPE$
              '
              'OrdeFabr.FindFirst "IdSubOr = '" & !IdSubOr & "' "
              SQLX$ = "SELECT * FROM OrdeFabr "
              SQLX$ = SQLX$ + " WHERE IdSubOr = '" & !idsubor & "' "
              Dim ORDERFFFABR As ADODB.Recordset
              Set ORDERFFFABR = FLEXCONN.Execute(FILTSQL$(SQLX$))
              With ORDERFFFABR
                  On Error Resume Next
                  .MoveFirst
                  'If OrdeFabr.NoMatch = False Then
                  If Err < 1 Then
                    On Error GoTo 0
                    'On Error Resume Next
                    'TETABU1$(5, CAITAB1%) = OrdeFabr!Destino
                    TETABU1$(5, CAITAB1%) = !DESTINO
                  End If
                  On Error GoTo 0
              End With
              Set ORDERFFFABR = Nothing
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
              TETABU1$(11, CAITAB1%) = FORMATNUM$(!CANTREAL, "F5.0")
              TETABU1$(12, CAITAB1%) = FORMATNUM$(!CaNoConf, "F5.0")
              TETABU1$(13, CAITAB1%) = ""
              NOCON = 0
              If !CANTREAL > 0.01 Then
                 NOCON = (!CaNoConf / !CANTREAL) * 100
              End If
              TETABU1$(14, CAITAB1%) = FORMATNUM$(NOCON, "F5.1")
              CADEN = 0
              If TIENETO > 0 Then
                 TIEHOR = TIENETO / 60
                 CADEN = (!CANTREAL / TIEHOR)
              End If
              TETABU1$(15, CAITAB1%) = FORMATNUM$(CADEN, "F10.1")
           '
            End If
        .MoveNext
        Loop
        '
    End With
    Set PROXFECTMP = Nothing
      '
    Call PRINTDOC("FAUTCONT")
      '
      '27/03/07 (FIN)
      '
End Sub


Sub CamStatNull()
    '' pone en 0 los registros null del campo Status y RETRABA de tabla INFOFAB
    '
    Call ABREORFAB
    '
    CONDI$ = "STATUS IS NULL"
    Call ACTUREGISTRO("INFOFAB", "STATUS", CONDI$, 0, REGAF&)
    '
    CONDI$ = "RETRABA IS NULL"
    Call ACTUREGISTRO("INFOFAB", "RETRABA", CONDI$, 0, REGAF&)
    '

End Sub

Sub InformesPorMotivo()

    Call DESHASFECHA(Desde%, Hasta%, PERIODO$)
    If PERIODO$ <> "" Then
        Call SELECHO("TAMOTIV/Tabla de Motivos")
        Motivo% = XVALO(VALACT1$("TAMOTIV"))
        DescripcionMotivo$ = TRIM$(VALACT2$("TAMOTIV"))
        
        Campos$ = "Fecha/D8;O-Fab/A18;Código/A24;Descripción/A32;C.Op/I4;Operario/A24;C.Ope/I5;Operación/A20;C.Mq/F7.2;Máquina/A20;Cantidad/F10.1"
        If Motivo% > 0 Then
            Campos$ = Campos$ & ";C.Mt/I0;Motivo/A0"
        Else
            Campos$ = Campos$ & ";C.Mt/I4;Motivo/A20"
            DescripcionMotivo$ = "TODOS"
        End If
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        
        GRILLAGRAL.Caption = "INFORMES POR MOTIVO (" & DescripcionMotivo$ & ") | " & PERIODO$
        
        DesdeSql$ = FETEXCOR1$(Desde%)
        Hasta% = DIPOST(Hasta%)
        HastaSql$ = FETEXCOR1$(Hasta%)
        Dim da As New DatosAuxiliares
        IdTablaOperarios& = da.ObtenerIdTablaAuxiliar("TABOPERA")
        IdTablaMotivos& = da.ObtenerIdTablaAuxiliar("TAMOTIV")
        
        Dim rst As ADODB.Recordset
        FiltroMotivo$ = "motirecha=" & CStr(Motivo%)
        CondicionVinculo$ = " inner"
        If Motivo% = 0 Then
            FiltroMotivo$ = "motirecha>=0"
            VinculoJoin$ = " left"
        End If
        ConsultaSql$ = "select im.fehorlib, im.idsubor, im.cod_exte, im.descrip, im.operario, im.campo2, im.numeoper, im.descoper, im.maquina, im.descmaq, im.cantreal, im.motirecha, ds.campo2 from" _
            & " (select i.fehorlib, i.idsubor, i.cod_exte, i.descrip, i.operario, ds.campo2, i.numeoper, i.descoper, i.maquina, i.descmaq, i.cantreal, i.motirecha from" _
            & " (" _
            & " (select fehorlib, idsubor, cod_exte, descrip, ltrim(rtrim(substring(legaops,1,2))) operario, numeoper, descoper, maquina, descmaq, cantreal, motirecha" _
            & " from infofab with(nolock) where " & FiltroMotivo$ & " and fehorlib>='" & DesdeSql$ & "' and fehorlib<'" & HastaSql$ & "') i" _
            & " inner join datasql ds with(nolock) on (i.operario=ds.campo1))" _
            & " where ds.id_lista = " & CStr(IdTablaOperarios&) & ") im" _
            & VinculoJoin$ & " join " _
            & " (select * from datasql ds where ds.id_lista= " & CStr(IdTablaMotivos&) & ") ds" _
            & " on (im.motirecha=ds.campo1)" _
            & " order by im.fehorlib, im.numeoper"

        Set rst = READSET(ConsultaSql$)
        
        Call Carga_MSF_Recordset(rst, Campos$, GRILLAGRAL.GRID)
        
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.MnuEdición.Visible = False
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        
        GRILLAGRAL.Show 1
        
        Dim MSF As MSFlexGrid
        If GRILLAGRAL.Menu1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOPMO", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.Menu1 = 0
            Screen.MousePointer = 1
        End If
    End If

End Sub

Sub ReImpresionEtiquetaInformeOperacion()

    Dim ls As New ListaSeleccion
    
    ls.TITULO = "REIMPRESION DE ETIQUETA - LISTADO DE INFORMES"
    
    ls.EncabezadoDeColumnas = "O-Fabricacion/A18;Partida/A20;Código/A24;Descripción/A32;N-O/I3;Operación/A20;Cantidad/F10.1"
    
    ls.ConsultaSql = "select ifo.idsubor, ifo.idenlote, ifo.cod_exte, ifo.descrip, ifo.numeoper, s.descoper, ifo.cantreal from" _
    & " (" _
    & " select i.idsubor, o.idenlote, o.cod_inte, o.cod_exte, o.descrip, i.numeoper, i.cantreal from" _
    & " (select idsubor, numeoper, sum(cantreal) cantreal from infofab group by idsubor, numeoper) i" _
    & " inner join ordefabr o on (i.idsubor=o.idsubor)" _
    & " ) ifo" _
    & " inner join secoper s on (ifo.cod_inte=s.codiconj and ifo.numeoper=s.numeoper)" _
    & " order by ifo.idsubor, ifo.numeoper"
    
    ls.AnchoLista = 12000
    Call ls.Mostrar

    If ls.FilaElegida(1) <> "" Then
        OrdenFabricacion$ = ls.FilaElegida(1)
        NumeroOperacion% = CInt(ls.FilaElegida(5))
        
        Call ImpresionEtiquetaInformeOperacion(OrdenFabricacion$, NumeroOperacion%)
    End If
    
End Sub

Function TotalInformesAprobadosOperacionAnterior(OrdenFabricacion As String, NumeroOperacion As Integer) As Single

    Dim PrimeraOperacion As Boolean
    PrimeraOperacion = False
    
    CodigoInterno% = XVALO(VALOBADA("ordefabr", "cod_inte", "idsubor='" & OrdenFabricacion & "'"))
    OperacionAnterior% = XVALO(VALOBADA("secoper", "top 1 numeoper", "codiconj=" & CStr(CodigoInterno%) & " and numeoper<" & CStr(NumeroOperacion) & " order by numeoper desc"))
    If OperacionAnterior% = 0 Then
        PrimeraOperacion = True
        OperacionAnterior% = NumeroOperacion
    End If
    
    If PrimeraOperacion Then
        SALDO# = XVALO(VALOBADA("operfab", "cantprog", "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(OperacionAnterior%)))
    Else
        SALDO# = XVALO(VALOBADA("infofab", "sum(cantreal)", "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(OperacionAnterior%) & " and retraba<1 and motirecha=0"))
    End If
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
            
    TotalInformesAprobadosOperacionAnterior = SALDO#
    
End Function

Function PendienteDeInformarSegunOperacionAnterior(OrdenFabricacion As String, NumeroOperacion As Integer) As Single

    Dim PrimeraOperacion As Boolean
    PrimeraOperacion = False
    
    CodigoInterno% = XVALO(VALOBADA("ordefabr", "cod_inte", "idsubor='" & OrdenFabricacion & "'"))
    OperacionAnterior% = XVALO(VALOBADA("secoper", "top 1 numeoper", "codiconj=" & CStr(CodigoInterno%) & " and numeoper<" & CStr(NumeroOperacion) & " order by numeoper desc"))
    If OperacionAnterior% = 0 Then OperacionAnterior% = NumeroOperacion
    
    CantidadProgramadaOperacionAnterior# = XVALO(VALOBADA("operfab", "cantprog", "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(OperacionAnterior%)))
            
    CantidadRechazadaOperacionAnterior# = XVALO(VALOBADA("infofab", "sum(cantreal)", "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(OperacionAnterior%) & " and motirecha>0"))
    
    CantidadRealizadaOperacionActual# = XVALO(VALOBADA("infofab", "sum(cantreal)", "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(NumeroOperacion%) & " and retraba<1"))
            
    DisponibleOperacionAnterior# = (CantidadProgramadaOperacionAnterior# - CantidadRechazadaOperacionAnterior#)
    PendienteDeInformarSegunOperacionAnterior = DisponibleOperacionAnterior# - CantidadRealizadaOperacionActual#
    
End Function

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
    '
    If EXISTELREP%("INFPROXMAQ") = 1 Then
        Call FINAL("?INFPROXMAQ")
        GoTo 99
    End If
    '
    HOII% = HOY(HOS$)
    VDEF$ = MKI$(COMESACT%) + MKI$(HOII%) + Space$(6)
10  OBX$ = OBJPLA$("DESHASFECHMAQ", VDEF$)
    If OBX$ = "" Then GoTo 99
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Has% < Des% Then GoTo 10
    MAQUISEL$ = TRIM$(Mid$(OBX$, 5, 6))
    If MAQUISEL$ = "" Then GoTo 10
    If ECOARCH$("PADMAQ", MAQUISEL$) = "" Then GoTo 10
    '
    FEC1 = FETEXCOR1$(Des%)
    FEC2 = FETEXCOR1$(Has%, "HASTA") ' CON HASTA PARA QUE CONCATENE HORAS
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab "
    SQLX$ = SQLX$ + " WHERE  FeHorLib  >=  '" & FEC1 & "' "
    SQLX$ = SQLX$ + " AND FeHorLib <= '" & FEC2 & "' "
    SQLX$ = SQLX$ + " AND Maquina = '" & MAQUISEL$ & "' "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    'Set PROMAQTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PROMAQTMP As ADODB.Recordset
    Set PROMAQTMP = New ADODB.Recordset
    PROMAQTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
          Call REPLA(XX$, !idsubor, 3, 12)
          Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            DOPE$ = TRIM$(Str$(!NUMEOPER))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DESCOPER
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, !MAQUINA, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 73, 2)
          Call REPLA(XX$, !idsubor, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NUMEOPER), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CANTREAL), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf)
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
    Set PROMAQTMP = Nothing
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
   OPXX% = ALTERNA%("Configuracion de Formulario\Maestro de Empleados")
   If OPXX% = 1 Then
      If DERACCE%("STPOFAB", "Configuracion de Formularios de Informes de Produccion") = 2 Then
          CFGINFPR07.Show 1
      End If
   ElseIf OPXX% = 2 Then
         GoTo 10
   End If
   '
10 OPXX% = ALTERNA%("Padron de Operarios\Células de Producción\Integrantes de Células")
   If OPXX% = 1 Then
        'Call FINAL("*PADRON")
        'AppActivate Caption
        'SendKeys Chr$(1)
        EMP_GES05.Show 1
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
    Call CONECRUN("ORFABR07", "Ordenes de Fabricación")
    Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    ACONEC$ = "CARMAQ07"
    Call CONECRUN(ACONEC$, "")
    Command25.Enabled = True
End Sub

Sub Command26_Click()
    Command26.Enabled = False
    Dim LEGAI%(2048), TIEREAL(2048), TIEMSTA(2048), RENDPROM(2048)
    Dim REGMIN&(2048), VAMIN(2048), REGMAX&(2048), VAMAX(2048)
    '
    CANOPERS% = 0
    RENDIMA = XVALO(Control("PADIAMOD", "RENDIMAX"))
    If RENDIMA < 1 Then RENDIMA = 99999
    '
    'ACA AGREGO
    ALT% = 0
    '27/03/07 (OT-07-0087)
    ALTMAX% = 3
    OPC$ = "Listados Generales\Por Operario\Por Pieza y Operación"
    If Control$("IFABRICA", "FAUTCONT") <> "" Then
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
      FEC1 = FETEXCOR1$(DES1%)
      FEC2 = FETEXCOR1$(HAS1%, "HASTA")
      
      '
      Call ABREORFAB
      SQLX$ = " SELECT * FROM InfoFab "
      SQLX$ = SQLX$ + " WHERE  FeHorLib  >=  '" & FEC1 & "' "
      SQLX$ = SQLX$ + " AND FeHorLib  <=  '" & FEC2 & "' "
      SQLX$ = SQLX$ + " ORDER BY IdSubOr "
      On Error Resume Next
      'Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim PROXFECTMP As ADODB.Recordset
      Set PROXFECTMP = New ADODB.Recordset
      PROXFECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
          Call REPLA(XX$, !idsubor, 3, 12)
          Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            DOPE$ = TRIM$(Str$(!NUMEOPER))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DESCOPER
          Call REPLA(XX$, DOPE$, 17, 42)
          On Error Resume Next
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 59, 2)
          Call REPLA(XX$, MKI$(!Hora_Des), 61, 2)
          Call REPLA(XX$, MKI$(!Hora_Has), 63, 2)
          On Error GoTo 0
          Call REPLA(XX$, !MAQUINA, 65, 6)
            HHORA$ = Format(!FeHorLib, "hh:mm:ss")
            If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then
               On Error Resume Next
               HHORA$ = HORATEXNOR$(!Hora_Has)
               On Error GoTo 0
            End If
            FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 73, 2)
            If IsNull(!Hora_Has) = True Then
              Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 63, 2)
            End If
          Call REPLA(XX$, !idsubor, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NUMEOPER), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CANTREAL), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf, ALTOP%)
          Call REPLA(XX$, MKS$(TSTAN), 155, 4)
          '
          ' ENJUAGUE ESCORIAL
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "ESCORIAL") > 0 Or InStr(EPAC$, "CONOMETAL") > 0 Then
             'TSTANBA = TIESTAN(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
             TSTANBA = TIEMAQUI(!Cod_Inte, !NUMEOPER, 1, 1, ALTOP%)
             TSTAN = 60 * TIEMAQUI(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf, ALTOP%)
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
      Set PROXFECTMP = Nothing
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
    RENDIMA = XVALO(Control("PADIAMOD", "RENDIMAX"))
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
    FEC1 = FETEXCOR1$(DES1%)
    FEC2 = FETEXCOR1$(HAS1%, "HASTA")
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab "
    SQLX$ = SQLX$ + " WHERE FeHorLib  >=  '" & FEC1 & "' "
    SQLX$ = SQLX$ + " AND  FeHorLib  <=  '" & FEC2 & "' "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PROXFECTMP As ADODB.Recordset
    Set PROXFECTMP = New ADODB.Recordset
    PROXFECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
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
              Call REPLA(XX$, !idsubor, 3, 12)
              Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
                DOPE$ = TRIM$(Str$(!NUMEOPER))
                If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
                DOPE$ = TRIM$(AJUSTI$(DOPE$, 7) + !DESCOPER)
                If ATMUL% > 0 Then DOPE$ = DOPE$ + "-AM"
              Call REPLA(XX$, DOPE$, 17, 42)
              On Error Resume Next
              Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 59, 2)
              Call REPLA(XX$, MKI$(!Hora_Des), 61, 2)
              Call REPLA(XX$, MKI$(!Hora_Has), 63, 2)
              On Error GoTo 0
              Call REPLA(XX$, !MAQUINA, 65, 6)
                HHORA$ = Format(!FeHorLib, "hh:mm:ss")
                If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
                FFECH$ = Format(!FeHorLib, "dd/mm/yy")
              Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
              Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 73, 2)
                If IsNull(!Hora_Has) = True Then
                  Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 63, 2)
                End If
              Call REPLA(XX$, !idsubor, 75, 12)
              Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
              Call REPLA(XX$, MKI$(!NUMEOPER), 89, 2)
              Call REPLA(XX$, DOPE$, 91, 48)
              Call REPLA(XX$, MKD$(!CANTREAL), 139, 8)
              Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
              TSTAN = 60 * TIESTAN(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf, ALTOP%)
              Call REPLA(XX$, MKS$(TSTAN), 155, 4)
              '
              ' ENJUAGUE ESCORIAL
              EPAC$ = TRIM$(UCase$(EMPREAC$))
              If InStr(EPAC$, "ESCORIAL") > 0 Or InStr(EPAC$, "CONOMETAL") > 0 Then
                 'TSTANBA = TIESTAN(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
                 TSTANBA = TIEMAQUI(!Cod_Inte, !NUMEOPER, 1, 1, ALTOP%)
                 TSTAN = 60 * TIEMAQUI(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf, ALTOP%)
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
    Set PROXFECTMP = Nothing
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
    Screen.MousePointer = 11
    '
    If EXISTELREP%("INFOPOROF") = 1 Then
        Call FINAL("?INFOPOROF")
        GoTo 99
    End If
    '
    'NUEVO
    Call ABREORFAB
    '
'    SQLX$ = " SELECT * FROM ORDEFABR "
'    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
'    '
'    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim ORDEFABTMP As ADODB.Recordset
'    Set ORDEFABTMP = New ADODB.Recordset
'    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'    With ORDEFABTMP
'      On Error Resume Next
'      .MoveLast
'      If Err Then
'         Call MENSERR(24, "NO Existen Ordenes de Fabricación. Imposible Continuar la Carga.")
'         On Error GoTo 0
'         GoTo 99
'      End If
'      .MoveFirst
'      FIN& = .RecordCount
'      JJ& = 0: JJ1& = 0: IDANT$ = ""
'      Do While (.EOF = False)
'        JJ1& = JJ1& + 1
'        Call TRACE(20, JJ1&, FIN&)
'        If !IdSubOr <> IDANT$ Then
'           Y$ = REGBLAN("ECORDEF")
'           '
'           Call REPLA(Y$, !IdSubOr, 1, 12)
'           Call REPLA(Y$, !Referen, 13, 52)
'           '
'           JJ& = JJ& + 1
'           Call GRAREG("ECORDEF", Y$, JJ&)
'           IDANT$ = !IdSubOr
'        End If
'        '
'      .MoveNext
'      Loop
'    End With
'    Screen.MousePointer = 1
'
'    Set ORDEFABTMP = Nothing
'    '
'    Call SETULTREG("ECORDEF", JJ&)
'    Call CIERRARCH("ECORDEF")
'    'HASTA ACA NUEVO
'10  Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación.")
'    NUORDFAB$ = TRIM$(VALACT1$("ECORDEF"))
     CondicionSql$ = "IDSUBOR <> '' ORDER BY IDSUBOR"
     Call CARGALISEL("LISTOTOF", "ORDEFABR", "IDSUBOR/A12; REFEREN/A48; ID_ORFA/A16", CondicionSql$)

10  Call SELECHO("LISTOTOF/Indice General de Ordenes de Fabricación.")
    NUORDFAB$ = TRIM$(VALACT1$("LISTOTOF"))

    '
11  NUORDFAB$ = AJUSTI$(VALACT1$("LISTOTOF"), 12)
    If TRIM$(NUORDFAB$) = "" Then
       Command27.Enabled = True
       Call FRESHFORM(Caption)
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab "
    SQLX$ = SQLX$ + " WHERE  IdSubOr  =  '" & NUORDFAB$ & "' "
    SQLX$ = SQLX$ + " ORDER BY NumeOper "
    'Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PROXFECTMP As ADODB.Recordset
    Set PROXFECTMP = New ADODB.Recordset
    PROXFECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
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
          Call REPLA(XX$, !idsubor, 3, 12)
          Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            DOPE$ = TRIM$(Str$(!NUMEOPER))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DESCOPER
          Call REPLA(XX$, DOPE$, 17, 48)
          Call REPLA(XX$, !MAQUINA, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 73, 2)
          Call REPLA(XX$, !idsubor, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NUMEOPER), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CANTREAL), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf)
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
    Set PROXFECTMP = Nothing
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
    '
    If EXISTELREP%("INFOPRODCOD") = 1 Then
        Call FINAL("?INFOPRODCOD")
        GoTo 99
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    VDEF$ = MKI$(COMESACT%) + MKI$(HOII%) + Space$(6)
10  OBX$ = OBJPLA$("DESHASFECHCOD", VDEF$)
    If OBX$ = "" Then GoTo 99
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Has% < Des% Then GoTo 10
    CIIX% = CVI(Mid$(OBX$, 5, 2))
    If CIIX% < 1 Then GoTo 10
    FEC1 = FETEXCOR1$(Des%)
    FEC2 = FETEXCOR1$(Has%, "HASTA") ' CON HASTA PARA QUE CONCATENE HORAS
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab"
    SQLX$ = SQLX$ + " WHERE FeHorLib  >=  '" & FEC1 & "' "
    SQLX$ = SQLX$ + " AND FeHorLib <= '" & FEC2 & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CIIX% & " "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    'Set PROMAQTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PROMAQTMP As ADODB.Recordset
    Set PROMAQTMP = New ADODB.Recordset
    PROMAQTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
          Call REPLA(XX$, !idsubor, 3, 12)
          Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            DOPE$ = TRIM$(Str$(!NUMEOPER))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DESCOPER
          Call REPLA(XX$, !DESCOPER, 17, 48)
          Call REPLA(XX$, !MAQUINA, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
          Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
          Call REPLA(XX$, MKI$(Int(HORANUM(HHORA$))), 73, 2)
          Call REPLA(XX$, !idsubor, 75, 12)
          Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
          Call REPLA(XX$, MKI$(!NUMEOPER), 89, 2)
          Call REPLA(XX$, DOPE$, 91, 48)
          Call REPLA(XX$, MKD$(!CANTREAL), 139, 8)
          Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
          TSTAN = 60 * TIESTAN(!Cod_Inte, !NUMEOPER, !CANTREAL, !NuOrInf)
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
    Set PROMAQTMP = Nothing
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
10 TTXX$ = "Los Listados Presentan las Operaciones Informadas"
   TTXX$ = TTXX$ + "\para las Ordenes de Fabricación en Curso de Ejecución."
   TTXX$ = TTXX$ + "\   \Opciones de Detalle de Presentación:"
   TTXX$ = TTXX$ + "\\Detalle por Informe\Acumulado por Operación\Saldos de Productos en Proceso"
   OPXX% = COMALTER%(TTXX$)
   '
   If OPXX% = 1 Then
        Call CALCUPROCE0(OPXX%)
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
    Call CONECRUN("EFIMOB07", "Control de Productividad de Mano de Obra")
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
   AJUPROCE07.Show 1        ' AJUSTE DE SALDOS PRODUCTOS EN PROCESO
   Command33.Enabled = True
End Sub

Private Sub Command34_Click()
    '\\\OT-08-00850-OD-16/08/08///
    Command34.Enabled = False
    Screen.MousePointer = 11
    
    Call ABREORFAB
    SQLX$ = "SELECT * FROM OPERFAB"

    Dim OPERTMP As ADODB.Recordset
    Set OPERTMP = New ADODB.Recordset
    OPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    FIN& = ULTREG&("OPERFAB")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      TTXX1$ = REGLEIDO$("OPERFAB", i&)
      NumeroOrFab$ = Mid$(TTXX1$, 1, 12)
      CodPieza% = CVI(Mid$(TTXX1$, 13, 2))
      DescriPieza$ = DESCRIT$(CodPieza%)
      CodOper% = CVI(Mid$(TTXX1$, 15, 2))
      DescOperac$ = Mid$(TTXX1$, 17, 48)
      CodOrFab% = CVI(Mid$(TTXX1$, 65, 2))
      MAQUINA$ = Mid$(TTXX1$, 151, 6)
      CantProgramada# = CVD(Mid$(TTXX1$, 79, 8))
      CantRealiz# = CVD(Mid$(TTXX1$, 87, 8))
      SaldoARealiz# = CVD(Mid$(TTXX1$, 95, 8))
      TiempPreparac& = CVS(Mid$(TTXX1$, 103, 4))
      TiempEjecuci& = CVS(Mid$(TTXX1$, 107, 4))
      TiempVarEjecu& = CVS(Mid$(TTXX1$, 111, 4))
      TiempDemDistr& = CVS(Mid$(TTXX1$, 115, 4))
      TiempMantHerr& = CVS(Mid$(TTXX1$, 119, 4))
      TiemLimp& = CVS(Mid$(TTXX1$, 123, 4))
      OtrosTiemp& = CVS(Mid$(TTXX1$, 127, 4))
      TiempTotoa& = CVS(Mid$(TTXX1$, 131, 4))
      FechaComProy = CVDAT(CVI(Mid$(TTXX1$, 135, 2)))
      HoraComProyect = CVI(Mid$(TTXX1$, 137, 2))
      FechaTerProy = CVDAT(CVI(Mid$(TTXX1$, 139, 2)))
      HoraTermProy = CVI(Mid$(TTXX1$, 141, 2))
      FechaTerTemp = CVDAT(CVI(Mid$(TTXX1$, 143, 2)))
      HoraTermTemp = CVI(Mid$(TTXX1$, 145, 2))
      CantidOperar# = CVS(Mid$(TTXX1$, 147, 4))
      Observaciones$ = Mid$(TTXX1$, 151, 48)
      TiempoEjecucion# = CVS(Mid$(TTXX1$, 161, 4))
      RegisSecoper& = CVS(Mid$(TTXX1$, 199, 4))
      EquipAux$ = Mid$(TTXX1$, 203, 18)
      ObserCorta$ = Mid$(TTXX1$, 225, 6)
      DescriMaq$ = ECOARCH("PADMAQ", MAQUINA$)
      '
      With OPERTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !idsubor = NumeroOrFab$
        !Cod_Inte = CodPieza%
        !Cod_Exte = CODEXT(CodPieza%)
        !Despcrip = Left(DescriPieza$, 64)
        !NUMEOPER = CodOper%
        !DESCOPER = DescOperac$
        !MAQUINA = MAQUINA$
        !DescMaq = DescriMaq$
        !EquiAux = EquipAux$
        !CantiOps = CantidOperar#
        !CANTPROG = CantProgramada#
        !CANTREAL = CantRealiz#
        !CANTPEND = SaldoARealiz#
        !TiemPrep = TiempPreparac&
        !TiemFijo = TiempEjecuci&
        !TiemVari = TiempVarEjecu&
        !TiemDemo = TiempDemDistr&
        !TiemMHer = TiempMantHerr&
        !TiemLimp = TiemLimp&
        !TiemOtro = OtrosTiemp&
        !TiemTota = TiempTotoa&
        !TiemEjec = TiempoEjecucion#
        !TiemPend = 0
        If SaldoARealiz# > 0 Then
          !TiemPend = TIESTAN(CodPieza%, CodOper%, CantRealiz#, 1)
        End If
        If CantRealiz# < CantProgramada# Then !StatOper = 1
        If CantRealiz# >= CantProgramada# Then !StatOper = 3
        !ComiePro = FechaComProy
        !TermiPro = FechaTerProy
        !TermiTemp = FechaTerTemp
        !Observa = Observaciones$
        '
        .Update
      End With
    Next i&
    OPERTMP.Close
    Set OPERTMP = Nothing
    Call COMUNI("Migracion Completada ")
    Command34.Enabled = True
    Screen.MousePointer = 1

End Sub

Private Sub Command35_Click()

Screen.MousePointer = 11
    
    Call ABREORFAB
    SQLX$ = "SELECT * FROM INFOFAB"

    Dim OPERTMP As ADODB.Recordset
    Set OPERTMP = New ADODB.Recordset
    OPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    FIN& = ULTREG&("OPERFAB")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      TTXX1$ = REGLEIDO$("OPERFAB", i&)
      NumeroOrFab$ = Mid$(TTXX1$, 1, 12)
      CodPieza% = CVI(Mid$(TTXX1$, 13, 2))
      DescriPieza$ = DESCRIT$(CodPieza%)
      CodOper% = CVI(Mid$(TTXX1$, 15, 2))
      DescOperac$ = Mid$(TTXX1$, 17, 48)
      CodOrFab% = CVI(Mid$(TTXX1$, 65, 2))
      Makina$ = Mid$(TTXX1$, 73, 6)
      CantProgramada# = CVD(Mid$(TTXX1$, 79, 8))
      CantRealiz# = CVD(Mid$(TTXX1$, 87, 8))
      SaldoARealiz# = CVD(Mid$(TTXX1$, 95, 8))
      TiempPreparac& = CVS(Mid$(TTXX1$, 103, 4))
      TiempEjecuci& = CVS(Mid$(TTXX1$, 107, 4))
      TiempVarEjecu& = CVS(Mid$(TTXX1$, 111, 4))
      TiempDemDistr& = CVS(Mid$(TTXX1$, 115, 4))
      TiempMantHerr& = CVS(Mid$(TTXX1$, 119, 4))
      TiemLimp& = CVS(Mid$(TTXX1$, 123, 4))
      OtrosTiemp& = CVS(Mid$(TTXX1$, 127, 4))
      TiempTotoa& = CVS(Mid$(TTXX1$, 131, 4))
      FechaComProy = CVDAT(CVI(Mid$(TTXX1$, 135, 2)))
      HoraComProyect = CVI(Mid$(TTXX1$, 137, 2))
      FechaTerProy = CVDAT(CVI(Mid$(TTXX1$, 157, 2)))
      HoraTermProy = CVI(Mid$(TTXX1$, 159, 2))
      FechaTerTemp = CVDAT(CVI(Mid$(TTXX1$, 143, 2)))
      HoraTermTemp = CVI(Mid$(TTXX1$, 145, 2))
      CantidOperar# = CVS(Mid$(TTXX1$, 147, 4))
      Observaciones$ = Mid$(TTXX1$, 151, 48)
      TiempoEjecucion# = CVS(Mid$(TTXX1$, 161, 4))
      LegajoEmp$ = Mid$(TTXX1$, 165, 8)
      TiempoStandar = CVS(Mid$(TTXX1$, 173, 4))
      RegisSecoper& = CVS(Mid$(TTXX1$, 199, 4))
      EquipAux$ = Mid$(TTXX1$, 203, 18)
      ObserCorta$ = Mid$(TTXX1$, 225, 6)
      DescriMaq$ = ECOARCH("PADMAQ", Makina$)
      '
      With OPERTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !idsubor = TRIM$(NumeroOrFab$)
            CCDD$ = TRIM$(CODEXT(CodPieza%))
            CAN% = Len(Str(CantProgramada#))
            DESCRI$ = CCDD$ & "-" & TRIM$(DESCRIT$(CodPieza%))
        !DescriOF = NumeroOrFab$ & Left(DESCRI$, 45 - CAN%) & " " & TRIM$(FORMATNUM(CantProgramada#, "F10.1"))
            NOI% = XVALO(VALOBADA("InfoFab", " MAX(NuOrInf)", "IdSubOr = '" & TRIM$(NumeroOrFab$) & " ' And NumeOper =" & CodOper%))
            NOI% = NOI% + 1
        !NuOrInf = NOI%
        !Cod_Inte = CodPieza%
        !Cod_Exte = CODEXT(CodPieza%)
        !Descrip = Left(DescriPieza$, 64)
        !NUMEOPER = CodOper%
        !AltOper = 0
        !DESCOPER = DescOperac$
        !RETRABA = 0
        !AtMulti = 0
        !canorpro = CantProgramada#
        !CANTREAL = CantRealiz#
        'If !CantReal <> 0 Then MsgBox !CantReal
        !CaNoConf = 0
        !MAQUINA = Makina$
        !DescMaq = DescriMaq$
        !LegaOps = LegajoEmp$
        !FeHorLib = FechaTerProy
        !Hora_Des = HoraComProyect
        !Hora_Has = HoraTermProy
        !Hora_Fri = 0
        !MaqLiber = ""
        !TiemEjec = TiempoEjecucion#
        !TiemStan = TiempoStandar
        RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
        If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = Left(Observaciones$, 30)
        !Referen = ""
        '
        .Update
      End With
    Next i&
    OPERTMP.Close
    Set OPERTMP = Nothing
    Call COMUNI("Migracion Completada ")
    Command34.Enabled = True
    Screen.MousePointer = 1

End Sub

Private Sub Command36_Click()
   Call FINAL("?INFOFABRI") ' REPORTE COMPLETO DE INFORMES DE PRODUCCION
End Sub
    
Private Sub Command37_Click()
   '
   Command37.Enabled = False
5  OPXX% = ALTERNA%("Maestro de Operarios\Configuracion de Formulario\Tabla de Motivo de Rechazos")
   If OPXX% = 2 Then
      If DERACCE%("STPOFAB", "Configuracion de Formularios de Informes de Produccion") = 2 Then
          CFGINFPR07.Show 1
          GoTo 5
      End If
   ElseIf OPXX% = 1 Then GoTo 10
   
   ElseIf OPXX% = 3 Then GoTo 50
   Else
    GoTo 99
   End If
   '
10 OPXX% = ALTERNA%("Padron de Operarios\Células de Producción\Integrantes de Células")
   If OPXX% = 1 Then
        'Call FINAL("*PADRON")
        'AppActivate Caption
        'SendKeys Chr$(1)
        EMP_GES05.Show 1
      ElseIf OPXX% = 2 Then
        Call DEFCELULAS
      ElseIf OPXX% = 3 Then
        Call CARCELULAS
      Else
        GoTo 5
   End If
   
   GoTo 10
   
50  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TAMOTIV", "TAMOTIV")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TAMOTIV/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TAMOTIV")
         X$ = REGLEIDO$("!TAMOTIV", U&)
         JJ& = JJ& + 1
         Call GRAREG("TAMOTIV", X$, JJ&)
      Next U&
      Call SETULTREG("TAMOTIV", JJ&)
      Call CIERRARCH("TAMOTIV")
    End If
    
    Dim da As New DatosAuxiliares
    Call da.CrearTablasAuxiliares
    idTabla& = da.ObtenerIdTablaAuxiliar("TAMOTIV")
    Call BORRAREGISTROS("datasql", "id_lista=" & CStr(idTabla&), REGAF&)
    
    For i& = 1 To ULTREG&("TAMOTIV")
        RegTamotiv$ = REGLEIDO$("TAMOTIV", i&)
        CodMotivo% = Asc(Mid$(RegTamotiv$, 1, 1))
        Denominacion$ = TRIM$(Mid$(RegTamotiv$, 2, 30))
        
        FLEXCONN.Execute "insert into datasql (id_lista, lista, codiempr, campo1, campo2)" _
        & " values (" & CStr(idTabla&) & ",'TAMOTIV',0,'" & CStr(CodMotivo%) & "','" & Denominacion$ & "')"
    Next i&
    
  End If
  
  GoTo 10
   
99 Command37.Enabled = True
   '
End Sub

Private Sub Command39_Click()
    ' RECUPERA DATOS DUPLICADOS GABAL
    Call ABRECONEXION
    Call CREACAMPO("", "INFOFAB", "TIEMEJANT", 6, 4)
    '
    ' PONE EN CERO EL DE LOS TIEMPOS ACUMULADOS
    Call ACTUREGISTRO("INFOFAB", "TIEMEJANT", "TIEMEJEC > 0", 0, RAFE&)
    '
    SQLX$ = "SELECT * FROM INFOFAB"
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = New ADODB.Recordset
    INFOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With INFOTMP
       Do While Not .EOF
          CIXI% = XVALO(!Cod_Inte)
          NOPE% = XVALO(!NUMEOPER)
          NORFA$ = SAFETEXT$(!idsubor)
          NORINF% = XVALO(!NuOrInf)
          TREG = XVALO(!TiemEjec)
          '
          SQLY$ = "COD_INTE = " & CIXI% & " "
          SQLY$ = SQLY$ + " AND NumeOper = " & NOPE% & ""
          SQLY$ = SQLY$ + " AND IdSuBor = '" & NORFA$ & "'"
          SQLY$ = SQLY$ + " AND NUORINF < " & NORINF%
          '
          TANTE = XVALO(VALOBADA("INFOFAB", "SUM(TIEMEJEC)", SQLY$))
          If TANTE >= TREG Then TANTE = 0
          '
          !TIEMEJANT = TANTE
          .Update
       .MoveNext
       Loop
    End With
    INFOTMP.Close
    Set INFOTMP = Nothing
    '
    Set INFOTMP = New ADODB.Recordset
    INFOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With INFOTMP
       Do While Not .EOF
          NOPE% = XVALO(!NUMEOPER)
          NORFA$ = SAFETEXT$(!idsubor)
          NORINF% = XVALO(!NuOrInf)
          TNUE = XVALO(!TiemEjec) - XVALO(!TIEMEJANT)
          !TiemEjec = TNUE
          .Update
       .MoveNext
       Loop
    End With
    INFOTMP.Close
    Set INFOTMP = Nothing
    '
    Call COMUNI("COMPLETO")
    '
End Sub

'\\\OT-08-00850-OD-FIN///

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
    If EXISTE%("APROYE07.EXE") < 1 Then
      Call MENSERR(24, "Modulo APROYE07 no Instalado\en la Presente Versión.")
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
    If InStr(EPAC$, "ESCORIAL") > 0 Or InStr(EPAC$, "ALEAR") Or InStr(EPAC$, "CONOMETAL") Then
        PADIAMOD07.Show 1
      Else
        Call INFAGEN
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("MODINFOP", "Modificar Partes de Fabricacion") > 0 Then
      If Not (InStr(EMPREAC$, "ESCORIAL") > 0 Or InStr(EMPREAC$, "ALEAR") Or InStr(EMPREAC$, "CONOMETAL") > 0) Then
        Call INFOPER
      Else
        MODIPAMOD07.Show 1
      End If
   End If
   Command8.Enabled = True

End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMASTO"
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
   Openforms = DoEvents
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture13)

    If Mid$(APLINVO$, 1, 8) = "INFONUOF" Then
        OrdenFabricacion$ = TRIM$(Mid(APLINVO$, 9))
        Call INFAGEN(OrdenFabricacion$)
        Call FINAL("")
    End If

    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    Call CamStatNull ' pone en 0 los registros null del campo Status Y RETRABA de tabla INFOFAB
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("EFIMOB07.EXE") < 1 Then
      Command31.Enabled = False
      Label11.Enabled = False
    End If
    '
End Sub

'
Sub INFAGEN(Optional OrdenFabricacion$)      ' INFORME DE PRODUCCION - GENERAL
    '
    EsReteplast% = 0
    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
    
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "INFORMES DE PRODUCCION - GENERALES"

100 NROOF$ = ""
    If OrdenFabricacion$ <> "" Then
        NROOF$ = OrdenFabricacion$
        GoTo 1407
    End If
    '
    Call GENRESDAT
    '
    ModalidadRetrabajo% = 0
    OPX% = COMALTER%("Opciones de Trabajo:\\Solo O/F's Pendientes\Solo O/F's Cerradas\Informe de Re-Trabajos")
    If OPX% = 1 Then
          ASELE$ = "ECORDEP"
       ElseIf OPX% = 2 Then
          'ASELE$ = "ECORDEF"
          If DERACCE%("INFOFCE", "Informes Sobre O/F's Cerradas") < 2 Then Exit Sub
          INFOPOFCE.Show 1
          Exit Sub
       ElseIf OPX% = 3 Then
          ModalidadRetrabajo% = 1
       Else
          Exit Sub
    End If
     '
'1405 Call FRESHECHO(ASELE$)
'     If ULTREG&(ASELE$) < 1 Then
'         Call MENSERR(24, "No hay Ordenes de Fabricación Pendientes.")
'         Exit Sub
'     End If
'     '
'     Call SELECHO(ASELE$ + "/Ordenes de Trabajo en Curso")
'1410 NROOF$ = AJUSTI$(VALACT1$(ASELE$), 12)
'     If TRIM$(NROOF$) = "" Then
'       Exit Sub
'     End If
1405 CONDI$ = "StatuOrf < 3 ORDER BY IdSubOr"
     Call CARGALISEL("ECOOPFAB", "ORDEFABR", "IDSUBOR/A12;REFEREN/A48;ID_ORFA/A24", CONDI$)
     Call SELECHO("ECOOPFAB/Ordenes de Trabajo en Curso")
     NROOF$ = AJUSTI$(VALACT1$("ECOOPFAB"), 12)
     If TRIM$(NROOF$) = "" Then
        GoTo 100
     End If
     '
     Screen.MousePointer = 11
     '
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
1407 Call ABREORFAB
     SQLX$ = " SELECT * FROM OPERFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & TRIM$(NROOF$) & "' "
     If ModalidadRetrabajo% = 1 Then
        SQLX$ = SQLX$ + "AND StatOper = 10 "                     ' solo retrabajos"
     Else
        SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     End If
     SQLX$ = SQLX$ + " ORDER BY NumeOper "
     '
     'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim FABRTMP As ADODB.Recordset
     Set FABRTMP = New ADODB.Recordset
     FABRTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
          NOPI% = XVALO(!NUMEOPER)
          CAPROD# = !CANTPROG      'Cantidad Programada
          CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPI% & " AND ReTraba < 1  AND STATUS < 8"
          If ModalidadRetrabajo% = 1 Then CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPI% & " AND ReTraba = 1  AND STATUS < 8"
          CAREA# = XVALO(VALOBADA("INFOFAB", "SUM(CantReal)", CONDI$))
          
          SALDO# = CAPROD# - CAREA: If SALDO# < 0 Then SALDO# = 0
          If EsReteplast% = 1 Then
            SALDO# = TotalInformesAprobadosOperacionAnterior(NROOF$, NOPI%) - CAREA#
          End If
          '
          If SALDO# > 0.05 Or EsReteplast% = 1 Then                                              ' CON LA CANTIDAD ORIGINAL
1412        NOPE% = !NUMEOPER      'numero de operación
            DOPE$ = !DESCOPER      'Descripcion de la operacion
            MAQUI$ = !MAQUINA      'Máquina
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
     Set FABRTMP = Nothing
        
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     Call FRESHECHO("!OPERINF")
     '
     Screen.MousePointer = 1
     '
1415 Call SELECHO("!OPERINF/Operaciones Programadas para " + NROOF$)
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
        If OrdenFabricacion$ <> "" Then Exit Sub
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
         If OrdenFabricacion$ <> "" Then GoTo 1407
         GoTo 1407
     End If
     '
End Sub
'

'
'
'
Sub ImpresionEtiquetaInformeOperacion(OrdenFabricacion As String, NumeroOperacion As Integer)
    
    FORIPRE$ = "ETIPROCE"
    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then
        Call COMUNI("Falta Formulario de Impresión de Etiqueta")
        Exit Sub
    End If
    
    Call CREATABLA_ARCHIVOS_QR
    USUARIO% = USNBR% Mod 100
    
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Dim da As New DatosAuxiliares
    
    CondicionSqlOf$ = "idsubor='" & OrdenFabricacion & "'"
    CodigoInternoFabricacion% = XVALO(VALOBADA("ordefabr", "cod_inte", CondicionSqlOf$))
    Codigo$ = CODEXT$(CodigoInternoFabricacion%)
    Descripcion$ = DESCRIT0$(CodigoInternoFabricacion%)
    Partida$ = TRIM$(SAFETEXT$(VALOBADA("ordefabr", "idenlote", CondicionSqlOf$)))
    LOTE$ = ".": If Partida$ <> "" Then LOTE$ = Partida$
    Operacion$ = VALOBADA("secoper", "descoper", "codiconj=" & CStr(CodigoInternoFabricacion%) & " and numeoper=" & CStr(NumeroOperacion))
        CondicionSqlInfofab$ = "idsubor='" & OrdenFabricacion & "' and numeoper=" & CStr(NumeroOperacion) & " and retraba<1 and motirecha=0"
    CantidadInforme# = XVALO(VALOBADA("infofab", "sum(cantreal)", CondicionSqlInfofab$))
    FechaInforme% = HOY(HO$)
    '
    TEXTOQR$ = OrdenFabricacion & "|" & CStr(NumeroOperacion)
    IMAG$ = "IMAGRECE_QR.JPG"
    IDENTIFICADOR$ = "QR_RECEPCION"
    ANCHOX% = 0
    ALTOX% = 0
    '
    On Error Resume Next
    Kill LUCOM$ & IMAG$
    On Error GoTo 0
    QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
    Call T_Espera(3)
    Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
    Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)

    FILTX$ = Codigo$ & ";" & Descripcion$ & ";" & LOTE$ & ";" & Operacion$ & ";.;" & FECHATEX$(FechaInforme%) & ";" & TRIM$(FORMATNUM$(CantidadInforme#, "F10.1")) & ";" & OrdenFabricacion$
    Call STDFORM(FORIPRE$, FILTX$, "PRINT")
    
    If PFAC$ <> "" Then Call SETPRINPORT("RESTORE")

End Sub

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
     j% = 0
     MENSAIN$ = "Inconsistencia en Base de Datos de Ope-\raciones de Fabricación.\  \Los Punteros de  Operaciones Pendientes\no se Corresponden con los Registros de\Operaciones Programadas.\  \Consulte con su Soporte de Sistemas."
     Call FRESHECHO("!OPERINF")
     For k% = 1 To Len(RGFSX$) Step 4
        REGI& = CVS(Mid$(RGFSX$, k%, 4))
        XX$ = REGLEIDO$("OPERFAP", REGI&)
        Des% = CVI(Mid$(XX$, 135, 2))
        Has% = CVI(Mid$(XX$, 139, 2))
        If DIA% >= Des% Then
           If DIA% <= Has% Then
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
                 j% = j% + 1
                 YY$ = MKI$(j%) + Space$(58) + MKS$(REGI&)
                 NOPX$ = AJUSTI$(NROOF$, 13) + AJUSTI$(DESTI$, 9)
                 NOPX$ = NOPX$ + AJUSTD$(Str$(NOPE%), 5) + " - " + AJUSTI$(DOPE$, 38)
                 NOPX$ = Left$(NOPX$, 46)
                 NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
                 Call REPLA(YY$, NOPX$, 3, 58)
                 JJ& = j%
                 Call GRAREG("!OPERINF", YY$, JJ&)
                 '
              End If
           End If
        End If
     Next k%
     '
     If j% < 1 Then
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
     '\\\OT-08-0652-OD-11/09/08///
     Call ACTECOR(1)
     '\\\OT-08-0652-OD-FIN///
'     Call FRESHECHO("ECORDEF")
'     If ULTREG&("ECORDEF") < 1 Then
'         Call MENSERR(24, "No hay Ordenes de Fabricación en Curso.")
'         Exit Sub
'     End If
'     '
'1405 Call SELECHO("ECORDEF/Ordenes de Trabajo en Curso")
'     Call FRESHALL
'     '
'1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEF"), 12)
'     If TRIM$(NROOF$) = "" Then
'         Exit Sub
'     End If

     CondicionSql$ = "IDSUBOR <> '' ORDER BY IDSUBOR"
1405 Call CARGALISEL("LISTOTOF", "ORDEFABR", "IDSUBOR/A12; REFEREN/A48; ID_ORFA/A16", CondicionSql$)
    
1410 Call SELECHO("LISTOTOF/Indice General de Ordenes de Fabricación.")
     
     NROOF$ = TRIM$(VALACT1$("LISTOTOF"))
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     'ACA GENERAR CONSULTA CONTRA TABLA SQL EN OPERFAB
'     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
'     RFF$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
'     JJ& = 0
'     For KK% = 1 To Len(RFF$) Step 4
'       REOPER& = CVS(Mid$(RFF$, KK%, 4))
'       ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
'       NORFA$ = Left$(ZZ$, 12)
'       If NORFA$ = NROOF$ Then
'         CAREA# = CVD(Mid$(ZZ$, 87, 8))
      JJ& = 0
      SQLX$ = "SELECT * FROM INFOFAB"
      SQLX$ = SQLX$ + " WHERE IdSuBor = '" & NROOF$ & "'"
      Dim OPERFTMP As ADODB.Recordset
      Set OPERFTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With OPERFTMP
        If (.EOF And .BOF) Then
          Call MENSERR(24, "No Hay Operaciones Informadas\para la " + NROOF$)
          GoTo 1405
        End If
        STA% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IdSuBor = '" & NROOF$ & "'", "NOMESS"))
        
        Do While Not .EOF
          MAQUI$ = SAFETEXT(!MAQUINA)
          TITOT = XVALO(!TiemEjec)
          CAREA# = XVALO(!CANTREAL)
          LEGOPER$ = SAFETEXT(!LegaOps)
          FERE% = CVINT(!FeHorLib)
          'HORAYMIN = SAFETEXT(REPLACAR(Mid$(!FeHorLib, 12, 5), ":", ""))
          HORE% = !Hora_Has 'XVALO(HORAYMIN)
          REOPER& = !NuOrInf
'          MotRech% = XVALO(!MOTIRECHA)
'          RETR% = XVALO(!RETRABA)
'          RETRABAJO$ = "N"
'          If RETR% > 0 Then RETRABAJO$ = "S"
          
'         If Mid$(ZZ$, 177, 6) <> String$(6, 0) Then ' solo las ya informadas
'            MAQUI$ = Mid$(ZZ$, 151, 6): If TRIM$(MAQUI$) = "" Then MAQUI$ = Mid$(ZZ$, 73, 6)
'            FERE% = CVI(Mid$(ZZ$, 157, 2))
'            HORE% = CVI(Mid$(ZZ$, 159, 2))
'            LEGOPER$ = Mid$(ZZ$, 165, 6)
'            LEGA4% = CVI(Mid$(ZZ$, 171, 2))
'            If LEGA4% > 0 Then
'              Call MENSERR(24, "Informe no Editable")
'              GoTo 1405
'            End If
'            TITOT = CVS(Mid$(ZZ$, 161, 4))
            HORAS% = Int(TITOT / 60)
            HOX = HORAS% ' SE MODIFICA PASANDO A UNA VARIABLE VARIANT
            'POR QUE APARENTEMENTE VISUAL CUANDO REALIZA UNA OPERACION LLEVA LAS VARIABLES
            'AL VALOR + CHICO Y EN ESTE CASO SI TITOT ES MAYOR A 32767 SE DESBORDA.
            MINUS% = Int(TITOT - (60 * HOX))
            SEGUS% = 60 * (TITOT - 60 * HOX - MINUS%)
            While SEGUS% > 59: SEGUS% = SEGUS% - 60: MINUS% = MINUS% + 1: Wend
            While MINUS% > 59: MINUS% = MINUS% - 60: HORAS% = HORAS% + 1: Wend
            XX$ = REGBLAN$("OPERFAIN")
            '           '
            Call REPLA(XX$, NROOF$, 1, 12)
            Call REPLA(XX$, MKI$(!Cod_Inte), 13, 2)
            Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            Call REPLA(XX$, SAFETEXT(!DESCOPER), 17, 48)
            Call REPLA(XX$, MKD$(CAREA#), 65, 8)
            Call REPLA(XX$, MAQUI$, 73, 6)
            Call REPLA(XX$, MKI$(FERE%), 79, 2)
            Call REPLA(XX$, MKI$(HORE%), 81, 2)
            Call REPLA(XX$, MKI$(HORAS%) + Chr$(MINUS%) + Chr$(SEGUS%), 84, 4)
            Call REPLA(XX$, LEGOPER$, 118, 6)
'            Call REPLA(XX$, RETRABAJO$, 129, 1)
'            Call REPLA(XX$, Chr$(MotRech%), 130, 1)
            Call REPLA(XX$, MKS$(REOPER&), 125, 4)
            '
            JJ& = JJ& + 1

            Call GRAREG("!OPERFAIN", XX$, JJ&)
            '
        .MoveNext
        Loop
      End With
      OPERFTMP.Close
      Set OPERFTMP = Nothing
      Call SETULTREG("!OPERFAIN", JJ&)
     '
'     If JJ& < 1 Then
'        Call MENSERR(24, "No Hay Operaciones Informadas\para la " + NROOF$)
'        GoTo 1405
'     End If
     '
     TITUPA$ = VALACT2$("LISTOTOF")
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
     ATRI$ = "/POSEDIT(14)=2"
     VTB% = VENTABU%("INFOPER/TITUPAN=Informe de Produccion: " + ATRI$)
     '
     If VTB% >= 0 Then
        '
'        While Len(RFF$) < 4 * ULTREG&("!OPERFAIN")
'           RFF$ = RFF$ + Chr$(0)
'        Wend
        '
        U& = 0
        For U& = 1 To ULTREG&("!OPERFAIN")
           '
           XX$ = REGLEIDO$("!OPERFAIN", U&)
           HMS$ = Mid$(XX$, 84, 4)
           TITOT = 60 * CVI(Left$(HMS$, 2)) + Asc(Mid$(HMS$, 3, 1)) + Asc(Mid$(HMS$, 4, 1)) / 60
           Call REPLA(XX$, MKS$(TITOT), 84, 4)
           Call REPLA(XX$, NROOF$, 1, 12)
           Call REPLA(XX$, MKI$(CIPIE%), 13, 2)
           REOPER& = CVS(Mid$(XX$, 125, 4)) 'NUMERO DE ORDEN DE OPERACIONES
           NUMEOPE% = CVI(Mid$(XX$, 15, 2))
           '
           ' ACTUALIZA TABLA INFOFAB
           '
           SQLX$ = "SELECT * FROM INFOFAB"
           SQLX$ = SQLX$ + " WHERE IdSuBor = '" & NROOF$ & "'"
           SQLX$ = SQLX$ + " AND NuOrInf = " & REOPER& & ""
           SQLX$ = SQLX$ + " AND NumeOper = " & NUMEOPE% & ""
           Dim INFOTMP As ADODB.Recordset
           Set INFOTMP = New ADODB.Recordset
           INFOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With INFOTMP
            On Error Resume Next
            .MoveFirst
             If Err < 1 Then
'            If REOPER& > 0 And REOPER& <= ULTREG&("OPERFAB") Then
'               ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
'               NORFA$ = Left$(ZZ$, 12)
'               If NORFA$ <> NROOF$ Then
'                   Call CIERRARCH("*.*")
'                   Call MENSERR(24, "Imposible Grabar")
'                   Call FINAL("")
'               End If
               CAREA# = CVD(Mid$(XX$, 65, 8))
               !CANTREAL = XVALO(CAREA#)
               'Call REPLA(ZZ$, MKD$(CAREA#), 87, 8)
               MAQUI$ = Mid$(XX$, 73, 6)
               !MAQUINA = SAFETEXT(MAQUI$)
               'Call REPLA(ZZ$, MAQUI$, 151, 6)
               FERE% = CVI(Mid$(XX$, 79, 2))
               !FeHorLib = CVDAT(FERE%)
               'Call REPLA(ZZ$, MKI$(FERE%), 157, 2)
               HORAA% = CVI(Mid$(XX$, 81, 2))
               'Call REPLA(ZZ$, MKI$(FERE%), 159, 2)
               !Hora_Has = HORAA%
               LEGOPER$ = Mid$(XX$, 118, 6)
               'Call REPLA(ZZ$, LEGOPER$, 165, 6)
               !LegaOps = SAFETEXT(LEGOPER$)
               'Call REPLA(ZZ$, MKS$(TITOT), 161, 4)
               !TiemEjec = TITOT
               '
               NOPE% = CVI(Mid$(XX$, 15, 2))
               !NUMEOPER = XVALO(NOPE%)
               !NuOrInf = XVALO(REOPER&)
               '
               TSTAN = TIESTAN(CIPIE%, NOPE%, CAREA#, NUOROP%)
               !TiemStan = TSTAN
'               RETRABAJO$ = Mid$(XX$, 129, 1)
'               RETR% = 0
'               If UCase$(RETRABAJO$) = "S" Then RETR% = "1"
'               !RETRABA = 0
'               !MOTIRECHA = Asc(Mid$(XX$, 130, 1))

               '
               .Update
             End If
           End With
           INFOTMP.Close
           Set INFOTMP = Nothing

           '
           '
           ' ACTUALIZA TABLA OPERFAB
           ' CORRIGE EL PENDIENTE DE LA OPERACION TENIENDO EN CUENTA
           ' QUE PUEDEN EXISTIR INFORMES PARCIALES
           '
           SQLX$ = "SELECT * FROM OPERFAB"
           SQLX$ = SQLX$ + " WHERE IdSuBor = '" & NROOF$ & "'"
           SQLX$ = SQLX$ + " AND NumeOper = " & NUMEOPE% & ""
           Dim OPERFABTMP As ADODB.Recordset
           Set OPERFABTMP = New ADODB.Recordset
           OPERFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With OPERFABTMP
              CanProg# = !CANTPROG
              SALRE# = CanProg# - CAREA#
              If SALRE# < 0 Then SALRE# = 0
              !CANTPEND = SALRE#
              .Update
           End With
           OPERFABTMP.Close
           Set OPERFABTMP = Nothing

'          Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
'          Call REPLA(ZZ$, MKS$(TSTAN), 173, 4)
'          '
'          Call GRAREG("OPERFAB", ZZ$, REOPER&)
        Next U&
     End If
     '

     'Call CIERRARCH("OPERFAB")
     GoTo 1405
     '
     '\\\OT-08-0652-OD-FIN///

End Sub
'
Sub CARGAINF(NROOF$, NOPE%)
     '
     EsReteplast% = 0
     If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
     '
     HOII% = HOY(HOS$)
     AGREGA% = 0
     '
     Call ABREORFAB
     '
     SQLX$ = " SELECT * FROM OPERFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     'Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim OPFABTMP As ADODB.Recordset
    Set OPFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
       DOPE$ = !DESCOPER
       CAPROG# = !CANTPROG
       CAREAL# = 0    'CAPROG#    ' !CantReal
       TIETOT = !TiemTota
       TIEPEN = !TiemPend
       
       MAQUI$ = !MAQUINA
       FEOP% = CVINT(!TermiPro)
       If FEOP% < 1 Then
           FEOP% = HOII%
           HOOP% = Int(HORANUM(Time$))
         Else
           FEOPX$ = Format$(!TermiPro, "dd/mm/yy  hh:mm:ss")
           FEOP% = FECHANUM(Left$(FEOPX$, 8))
           HOOP% = HORANUM(Mid$(FEOPX$, 11, 8))
       End If
       MALIB$ = "S"
       
       If XVALO(!StatOper) = 10 Then ModalidadRetrabajo% = 1
       
     End With
     Set OPFABTMP = Nothing
     '
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     If ModalidadRetrabajo% = 1 Then SQLX$ = SQLX$ + " AND ReTraba = 1"
     '
     CAINFANT% = 0
     'Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim INFOFABTMP As ADODB.Recordset
     Set INFOFABTMP = New ADODB.Recordset
     INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If XVALO(!RETRABA) < 1 Or ModalidadRetrabajo% = 1 Then          'la condicion no se agrega en la consulta para que asi acumule los numeros de informe
             If XVALO(!Status) < 8 Then
              CAREAL# = CAREAL# + !CANTREAL     'si es retrabajo no toma en cuenta los saldos
             End If
           End If
           TIEINF = TIEINF + !TiemEjec
           CAINFANT% = CAINFANT% + 1
         .MoveNext
         Loop
       End If
     End With
     Set INFOFABTMP = Nothing
     '
     SALDO# = CAPROG# - CAREAL# - Merma#
     If EsReteplast% = 1 Then
        SALDO# = TotalInformesAprobadosOperacionAnterior(NROOF$, NOPE%) - CAREAL#
     End If
     '
     CONDI$ = "IDSUBOR = '" & NROOF$ & "'AND NumeOper = " & NOPE% & ""
     Stat% = XVALO(VALOBADA("INFOFAB", "STATUS", CONDI$, "NOMESS"))
     If CAPROG# >= 0.05 Then TIEPEN = TIETOT * SALDO# / CAPROG#
     If SALDO# < 0 Then
       SALDO# = 0
       TIEPEN = 0
     End If
     HORAS% = Int(TIEPEN / 60)
     '
     HRS& = HORAS%            ' SE CAMBIO  EL TIPO DE VARIABLE PARA QUE NO DESBORDE
     MNS& = 60 * HRS&         '  Y  MINUS% OBTENGA COMO RESULTADO 0
     MINUS% = Int(TIEPEN - MNS&)
     'MINUS% = Int(TIEPEN - 60 * HORAS%)
     
     'SEGUS% = 60 * (TIEPEN - 60 * HORAS% - MINUS%)
     SEGUS% = 60 * (TIEPEN - MNS& - MINUS%)
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
     ' *** ESTO APLICA SOLO A REGALIA
     If InStr(1, EMPREAC$, "MECANIZADOS") > 0 Or InStr(1, EMPREAC$, "REGALIA") > 0 Then
        Call REPLA(YY$, MKI$(HOY(HO$)), 79, 2)
            HoraActual% = HORANUM(Format(Now, "hh:mm"))
        Call REPLA(YY$, MKI$(HoraActual%), 81, 2)
            NumeroOperario% = CVI(Mid$(REGLEIDO$("PADRON", 1), 1, 2))
        Call REPLA(YY$, MKI$(NumeroOperario%), 118, 2)
     End If
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
     '
     RETRAB$ = ""
     NUMMOTIV% = 0
     Stat% = 0
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     
     If ModalidadRetrabajo% = 1 Then
        If ModalidadRetrabajo% = 1 Then RETRAB$ = "S"
        GoTo 1430
     End If
             
     If InStr(EPAC$, "ALEJPA") > 0 Or InStr(EPAC$, "REGALIA") > 0 Or InStr(EPAC$, "MECANIZADOS") > 0 Or InStr(EPAC$, "EDR") > 0 Or EsReteplast% > 0 Then
        NUMMOTIV% = 0
        VDEF$ = ""
        Call REPLA(VDEF$, NROOF$, 1, 12)
        Call REPLA(VDEF$, MKI$(CI%), 13, 2)
        Call REPLA(VDEF$, "N", 15, 1) ' VALOR POR DEFAULT
        '
1425    VDEF$ = OBJPLA$("RETRAMO", VDEF$)
        If VDEF$ = "" Then GoTo 1420
        '
        If DERACCE%("GENESCRAP", "Informe de Produccion con Motivo Scrap") > 1 Then
            RETRAB$ = TRIM$(Mid$(VDEF$, 15, 1))
            NUMMOTIV% = Asc(Mid$(VDEF$, 16, 1))
            '
            If UCase$(RETRAB$) <> "N" And UCase$(RETRAB$) <> "S" Then
               Call MENSERR(24, "Solo Valido 'S'(Re-Trabajo) o 'N'(No Re-Trabajo)")
               GoTo 1425
            End If
            '
            If UCase$(RETRAB$) = "S" And (NUMMOTIV% < 1 Or ECOARCH$("TAMOTIV", Chr$(NUMMOTIV%)) = "") Then
               Call MENSERR(24, "Ingreso de Motivo de Rechazo No Válido")
               GoTo 1425
            End If
            '
            TX$ = FILTERGETRECORD$("TAMOTIV", 1, Chr$(NUMMOTIV%))
            If Asc(Mid$(TX$, 32, 1)) > 0 Then ' SI ESTA MARCADO ESTE MOTIVO EN LA COLUMNA DE SCRAP CON 1
               Stat% = 8
            End If
            '
            If UCase$(RETRAB$) = "N" And Stat% = 8 Then
               Call MENSERR(24, "Ingreso de Motivo de Scrap\Solo si es Retrabajo")
               GoTo 1425
            End If
            '
            If Stat% = 8 Then
               TTXX$ = "Ha Ingresado un Informe Con Motivo Scrap\Esto 'Elimina' Todos los Informes Ingresados\Hasta el Momento,para la Presente O/F.\¿Esta Seguro de  Continuar?\\No, Cancelar\Sí, Continuar"
               If COMALTER%(TTXX$) < 2 Then
                  GoTo 1425
               End If
            End If
        End If
     End If
     '
1430 OBSE$ = TRIM$(Mid$(XX$, 88, 30))
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
     INI% = 1
     LEGOPER$ = Mid$(XX$, 118, 8)
     ACUOPERTEX$ = ""
     For U& = 0 To 3
        OPER% = CVI(Mid$(LEGOPER$, INI%, 2))
        OPERTEX$ = Right("  " + SAFETEXT$(OPER%), 2)
        ACUOPERTEX$ = ACUOPERTEX$ + OPERTEX$
        INI% = INI% + 2
     Next U&

     '
     STATUOP% = 3 ' Cumplida
     If CAINF# < SALDO# - 0.05 Then
       STATUOP% = 1   ' AUN PENDIENTE
       If ModalidadRetrabajo% = 1 Then STATUOP% = 10
       TTXX$ = "Queda un Saldo Pendiente\para esta Operación\\Mantener Pendiente\Cancelar Saldo"
       If COMALTER%(TTXX$) = 2 Then
         STATUOP% = 8 ' CANCELADO
       End If
     End If
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM InfoFab "
     Dim INFOFFFAB As ADODB.Recordset
     Set INFOFFFAB = New ADODB.Recordset
     INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With InfoFab
     With INFOFFFAB
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !idsubor = NROOF$
        !DescriOF = Left(ECOARCH$("ECORDEF", NROOF$), 64)
        !NuOrInf = CAINFANT%
        !Cod_Inte = CI%
        !Cod_Exte = CODEXT$(CI%)
        !Descrip = Left(DESCRIT$(CI%), 64)
        !NUMEOPER = NOPE%
        !AltOper = 0
        MARETRA% = 0: If UCase$(RETRAB$) = "S" Then MARETRA% = 1
        !RETRABA = MARETRA%
        !DESCOPER = Left(DOPE$, 48)
        !canorpro = CAPROG#
        !CANTREAL = CAINF#
        !MAQUINA = MAQUI$
        !DescMaq = Left(ECOARCH$("PADMAQ", MAQUI$), 64)
'        !LegaOps = LEGOPER$    'NO ESTA GUARDANDO EL DATO ,VERIFICAR!!!
        !LegaOps = ACUOPERTEX$
        !FeHorLib = FEHOLI
        !Hora_Des = 0
        !Hora_Has = HOOP%
        !MaqLiber = MALIB$
        !TiemEjec = TIEINF
           Modo% = CAINFANT%
        !TiemStan = TIESTAN(CI%, NOPE%, CAINF#, Modo%)   ' 1 o mas es para que no considere preparacion
           'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
           ' A LOS CAMPOS DEL REGISTRO. "!"
           'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
           RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
           If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = OBSE$
        !Referen = " "
        !MOTIRECHA = NUMMOTIV%
        !Status = Stat%
        !Procesado = 0
       .Update
     End With
     'InfoFab.Close
     Set INFOFFFAB = Nothing
     '
     If Stat% = 8 Then
        CONDI$ = "IdSubOr = '" & NROOF$ & "'"
        Call ACTUREGISTRO("INFOFAB", "STATUS", CONDI$, Stat%, REGAF&, "NOMESS")
     End If

     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     If ModalidadRetrabajo% = 1 Then SQLX$ = SQLX$ + " AND ReTraba = 1"
     '
     CAREAL# = 0: TIEINF = 0
     'Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim INFOFABTMP1 As ADODB.Recordset
    Set INFOFABTMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With INFOFABTMP1
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If XVALO(!RETRABA) < 1 Or ModalidadRetrabajo% = 1 Then      'si es retrebajo, no toma en cuenta los saldos
             CAREAL# = CAREAL# + !CANTREAL
           End If
           TIEINF = TIEINF + !TiemEjec
         .MoveNext
         Loop
       End If
     End With
     Set INFOFABTMP = Nothing
     '
     'With OperFab
     '.FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
     SQLX$ = " SELECT * FROM OperFab"
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     If ModalidadRetrabajo% = 1 Then SQLX$ = SQLX$ + " AND StatOper = 10"
     Dim IDSSSUBOR As ADODB.Recordset
     Set IDSSSUBOR = New ADODB.Recordset
     IDSSSUBOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With IDSSSUBOR
        On Error Resume Next
        .MoveFirst
        'If .NoMatch = False Then
        If Err < 1 Then
          '.Edit
          !CANTREAL = CAREAL#
          !CANTPEND = !CANTPROG - CAREAL#
            If !CANTPEND < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CANTPEND = 0
            CANTPE# = !CANTPEND
            Modo% = CAINFANT%
          !TiemEjec = TIEINF
          !TiemPend = TIEMAQUI(CI%, NOPE%, CANTPE#, Modo%)  ' 1 es para que no considere preparacion
            If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
          !StatOper = STATUOP%
          .Update
        End If
     End With
     Set IDSSSUBOR = Nothing
     '
     If TRIM$(Control$("", "FORINPR")) <> "" Then
        CANCELPRINT% = 0 ' CONFIGRACION PARA IMPRESION FANTASMA
        CONPFA09.Show 1
        If CONPFA09.Label1 = "OK" Then
           If CONPFA09.Check8.Value = 1 Then CANCELPRINT% = 1 'CASO IMPRESION FANTASMA
        End If
        Call PRININPRO(NROOF$, NOPE%, CAINFANT%)
        CANCELPRINT% = 0
     End If
     '
     If EsReteplast% = 1 Then
        If PendienteDeInformarSegunOperacionAnterior(NROOF$, NOPE%) < 0.05 Then
            Call ImpresionEtiquetaInformeOperacion(NROOF$, NOPE%)
        End If
     End If
     
     '
End Sub
'

Sub GESALPROP(Modo%)
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
         Z$ = REGBLAN$("SPPROCE")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, NROOF$, 3, 12)
         Call REPLA(Z$, MKI$(CI%), 15, 2)
         Call REPLA(Z$, MKI$(NOPE%), 17, 2)
         Call REPLA(Z$, DOPE$, 19, 48)
         Call REPLA(Z$, MKD$(CAPROD#), 67, 8)
         W& = W& + 1
         Call GRAREG("SPPROCE", Z$, W&)
         NORFAX$(W&) = NROOF$: NOPEA% = NOPE%
         GoTo 19
         '
10       Z$ = REGLEIDO$("SPPROCE", KX&)
         NOPEA% = CVI(Mid$(Z$, 17, 2))
         If NOPE% < NOPEA% Then
            Call REPLA(Z$, MKI$(NOPE%), 17, 2)
            Call REPLA(Z$, DOPE$, 19, 48)
            Call GRAREG("SPPROCE", Z$, KX&)
         End If
      End If
19 Next U&
   '
   Call SETULTREG("SPPROCE", W&)
   Call CIERRARCH("SPPROCE")
   '
   If Modo% = 1 Then
         Call FILESORT("SPPROCE", "", 3, 4, "ASC")
         Call CONECRUN("*LISPPR1", "Saldos en Proceso por O.F")
         'Call CONECRUN("*SOSPPR1", "Saldos en Proceso por O.F")
       ElseIf Modo% = 2 Then
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





Private Sub mnuInformePorMotivo_Click()

    Call InformesPorMotivo

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
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With RESERTMP
      On Error Resume Next
      .MoveFirst
      'If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        Call REPLA(Y$, MKI$(XVALO(!Cod_Inte)), 1, 2)
        Call REPLA(Y$, SAFETEXT$(!idsubor), 3, 12)
        Call REPLA(Y$, MKI$(XVALO(!Cod_Inte)), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(XVALO(!CANTRES)), 19, 8)
        Call REPLA(Y$, MKD$(XVALO(!CANTCONS)), 27, 8)
        Call REPLA(Y$, MKD$(XVALO(!Saldo_Entre)), 35, 8)
        Call REPLA(Y$, SAFETEXT$(!IDENLOTE), 43, 16)
        Call REPLA(Y$, Chr$(XVALO(!Deposito)), 60, 1)
        Call REPLA(Y$, MKS$(XVALO(!USOUNIT)), 75, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("RESERVA", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Set RESERTMP = Nothing
    Call SETULTREG("RESERVA", JJ&)
    Call CIERRARCH("RESERVA")
    '
    Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.")   '  Imposible Continuar la Carga.")
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, !idsubor, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !idsubor), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Set ORDEFABTMP = Nothing
    Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
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
     SQLX$ = ""
     SQLX$ = "SELECT * FROM INFOPROY "
     SQLX$ = SQLX$ + " Order by Nuor_Inf ASC"
     'Set INFTMP = Ofabric.OpenRecordset(SQLX$, 4)
     
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
     HOOP% = Int(HORANUM(Time$))
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
     SQLX$ = " SELECT * FROM InfoProy "
     Dim INFOPPPROXY As ADODB.Recordset
     Set INFOPPPROXY = New ADODB.Recordset
     INFOPPPROXY.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

     With INFOPPPROXY
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
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
     Set INFOPPPROXY = Nothing
     'InfoProy.Close
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
  SQLX$ = "SELECT * FROM TAPROYEC"
  SQLX$ = SQLX$ + " WHERE STAT_PROY < 3 "
  Call ABREPLANCUEN
  'Set CIETMP = Contable.OpenRecordset(SQLX$, 4)
  Dim CIETMP As ADODB.Recordset
  Set CIETMP = New ADODB.Recordset
  CIETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
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
  Set CIETMP = Nothing
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
  SQLX$ = ""
  SQLX$ = "SELECT * FROM INFOPROY "
  SQLX$ = SQLX$ + " Order by Nuor_Inf ASC"
  'Set INFTMP = Ofabric.OpenRecordset(SQLX$, 4)
  Dim INFTMP As ADODB.Recordset
  Set INFTMP = New ADODB.Recordset
  INFTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
  Set INFTMP = Nothing
  '
5 Call SELECHO("INFAPROY")
  NUINF1& = XVALO(VALACT1$("INFAPROY"))
  If NUINF1& > 0 Then
    OPX1% = COMALTER%("¿ Confirma Anular el Informe: " + "IF." + TRIM$(Str$(NUINF1&)) + " ?\\No, Cancelar\Si, Cerrar")
    If OPX1% = 2 Then
      'With InfoProy
        '.FindFirst "Nuor_Inf = " & NUINF1& & ""
        SQLX$ = " DELETE FROM InfoProy "
        SQLX$ = SQLX$ + " WHERE Nuor_Inf = " & NUINF1& & ""
         FLEXCONN.Execute (FILTSQL$(SQLX$))
        'If .NoMatch = False Then
          '.Delete
        'End If
      'End With
      Call COMUNI("Anulación Completa")
    End If
  End If
  '
99 Screen.MousePointer = 1
End Sub
'\\\OT-06-0252-WAR-FIN///

Sub CALCUPROCE0(Optional OPC%)
     'PARA DESARROLLAR SALDOS DE PRODUCTOS EN PROCESO - 22/08/06
     '\\\OT-06-0335-22/08/06///
     Dim NOFA$() ' vector donde cargo ordenes fabricacion
     '
     VDEVU$ = ""
     If OPC% > 0 Then
        Call DESHASFECHA(DES1%, HAS1%, PERIO$)
        VDEVU$ = PERIO$
        If VDEVU$ <> "" Then
          FEC1 = FETEXCOR1$(DES1%)
          FEC2 = FETEXCOR1$(HAS1%, "HASTA")
        End If
     End If
     '
     Screen.MousePointer = 11
     '
     Call ABREORFAB ' funcion que abre bases de datos
     SQLX$ = " SELECT * FROM ordefabr WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE  statuorf < 3 "
     'Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
''     Set PROXFECTMP = New ADODB.Recordset
''     PROXFECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Dim PROXFECTMP As ADODB.Recordset
     Set PROXFECTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

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
          NOFA$(X1%) = TRIM$(!idsubor)
          'en la posicion X1% escribo "idsubor"
        .MoveNext ' avanzo al siguiente
        Loop
     End With
     Set PROXFECTMP = Nothing
     '
10   JJ& = 0 ' Tiene nº de registro en acuproce
     REBLA$ = REGBLAN$("ACUPROCE") '
     'crea un string de blancos como la long del archivo
     SQLX$ = "SELECT * FROM InfoFab WITH(NOLOCK) "
     If VDEVU$ <> "" Then 'SI SE SELECCIONO RANGO DE FECHAS
       SQLX$ = SQLX$ + " WHERE FEHORLIB >= '" & FEC1 & "'"
       SQLX$ = SQLX$ + " AND FEHORLIB <= '" & FEC2 & "'"
     End If
     
''     Set INFOFABTMP = New ADODB.Recordset
''     INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Dim INFOFABTMP As ADODB.Recordset
     Set INFOFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With INFOFABTMP
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
          If IsNull(!RETRABA) Or !RETRABA < 1 Then OBSEX$ = ""
          If TRIM$(!MAQUINA) = "Ajuste" Then OBSEX$ = "Ajuste"
          '
          NROOF$ = "" ' NRO ORDEN FABRICACION
          If IsNull(!idsubor) = False Then
             NROOF$ = TRIM$(!idsubor)
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
          Call REPLA(XX$, !idsubor, 3, 12)
          Call REPLA(XX$, MKI$(!NUMEOPER), 15, 2)
            DOPE$ = TRIM$(Str$(!NUMEOPER))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DESCOPER
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
          Call REPLA(XX$, MKD$(!CANTREAL), 129, 8)
          Call REPLA(XX$, OBSEX$, 169, 32)
          '
          JJ& = JJ& + 1 ' incrementa registro ACUPROCE
          Call GRAREG("ACUPROCE", XX$, JJ&)
          'Graba en acuproce, lo que hay en xx en pos jj
60      .MoveNext
        Loop
     End With
     Set INFOFABTMP = Nothing
     '
70   ' AGREGA ENTRADAS A STOCK POR CIERRE DE o/fABRIC
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM Movisto WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE  CodiMovi = 'OF' "
     If VDEVU$ <> "" Then 'SI SE SELECCIONO RANGO DE FECHAS
       SQLX$ = SQLX$ + " AND FECHMOV >= '" & FEC1 & "'"
       SQLX$ = SQLX$ + " AND FECHMOV <= '" & FEC2 & "'"
     End If
     '
     'Set Movisel = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
''     Set MOVISEL = New ADODB.Recordset
''     MOVISEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Dim MOVISEL As ADODB.Recordset
     Set MOVISEL = FLEXCONN.Execute(FILTSQL$(SQLX$))

     '
     With MOVISEL
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
          NROOF$ = TRIM$(!DoPriLar)
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
          Call REPLA(XX$, MKD$(!CANTINGRE), 129, 8)
          '
          JJ& = JJ& + 1 ' incrementa registro ACUPROCE
          Call GRAREG("ACUPROCE", XX$, JJ&)
          'Graba en acuproce, lo que hay en xx en pos jj
80      .MoveNext
        Loop
     End With
     Set MOVISEL = Nothing
     Screen.MousePointer = 1

     '
     
90   Call SETULTREG("ACUPROCE", JJ&) 'Pone marca de fin archivo
     'JJ& es el ultimo registro valido
     Call CIERRARCH("ACUPROCE")
     'cierra el archivo
     '
     'ordena
     '
     ' SI SE SELECCIONO RANGO DE FECHAS LO IMPRIME
     Call HEADERS("ACUPROCE", "")
     If VDEVU$ <> "" Then
        Call HEADERS("ACUPROCE", "Corresponde a: " + VDEVU$)
     End If
     '
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
                    TOTCANT# = CVD(Mid$(YY$, 129, 8)) + CANTU#
                    TOTHORAS = CVS(Mid$(YY$, 67, 4)) + HORAS
                    Call REPLA(YY$, MKD$(TOTCANT#), 129, 8)
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
    j& = 0
    TTXX$ = ""
    For i& = 1 To ULTREG&("!ICELULAS")
       X$ = REGLEIDO$("!ICELULAS", i&)
       If CVI(Left$(X$, 2)) > 0 Then
         IGRUPO% = CVI(Left$(X$, 2))
         For JJ& = 1 To Len(TTXX$) Step 2
           If CVI(Mid$(TTXX$, JJ&, 2)) = IGRUPO% Then GoTo 19
         Next JJ&
         j& = j& + 1
         Call GRAREG("ICELULAS", X$, j&)
         TTXX$ = TTXX$ + MKI$(IGRUPO%)
       End If
19  Next i&
    Call SETULTREG("ICELULAS", j&)
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
    j& = 0
    TTXX1$ = "": TTXX2$ = ""
    For i& = 1 To ULTREG&("!DCELULAS")
       X$ = REGLEIDO$("!DCELULAS", i&)
       If CVI(Mid$(X$, 3, 2)) > 0 Then
         NLEGA% = CVI(Mid$(X$, 3, 2))
         For JJ& = 1 To Len(TTXX1$) Step 2
           If CVI(Mid$(TTXX1$, JJ&, 2)) = NLEGA% Then GoTo 19
         Next JJ&
         Call REPLA(X$, MKI$(ICELU%), 1, 2)
         j& = j& + 1
         Call GRAREG("!DCELULAS", X$, j&)
         TTXX1$ = TTXX1$ + MKI$(NLEGA%)
         If TTXX2$ <> "" Then TTXX2$ = TTXX2$ + " / "
         TTXX2$ = TTXX2$ + TRIM$(Str$(NLEGA%))
       End If
19  Next i&
    Call SETULTREG("!DCELULAS", j&)
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


'
Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Call ABRECONEXION
   
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "INFOPROY"
           Call Command6_Click                 ' INFORME DE PROYECTO
       Case "INFOPERPEN"
           Call Command7_Click                 ' IINFORMES DE OPERACIONES PENDIENTES PO O/F
       Case "CORINFPRO"
           Call Command8_Click                 ' CORRECION DE INFORMES DE PRODUCCION
       Case "AJSALPROPRO"
           Call Command33_Click                ' AJUSTE DE SALDOS EN PRODUCTOS EN PROCESO
       Case "PADROPER"
           Call Command19_Click                ' PADRON DE OPERARIOS
       Case "PADREMPL"
           EMP_GES05.Show 1                    ' PADRON DE EMPLEADOS
       '
       'CONSULTAS Y LISTADOS
       '
       Case "INFPROFEC"
           Call Command26_Click                ' INFORME DE PRODUCCION POR FECHA
       Case "INFPROMAQ"
           Call Command15_Click                ' INFORME DE PRODUCCION POR MAQUINA
       Case "INFPRORFAB"
           Call Command27_Click                ' INFORME DE PRODUCCION POR O/F
       Case "INFPROCOD"
           Call Command28_Click                ' INFORME DE PRODUCCION POR CODIGO DE PRODUCTO
       Case "INFOPORF"
           Call Command4_Click                 ' OPERACIONES PENDIENTES POR ORDEN DE FABRICACION
       Case "INFOPORMAQ"
           Call Command5_Click                 ' OPERACIONES PENDIENTES POR MAQUINA/EQUIPO
       Case "COOPERPEN"
           Call Command2_Click                 ' CONTROL DE OPERACIONES PENDIENTES
       Case "COOPENPRO"
           Call Command3_Click                 ' CONTROL DE OPERACIONES Y PRODUCTOS EN PROCESO
       Case "INFOPMOT"
           Call InformesPorMotivo              ' INFORMES POR MOTIVO
       Case "REIMETI"
            Call ReImpresionEtiquetaInformeOperacion
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub PRININPRO(NORFA$, NUOPE%, NUMORINF%)
    '
    HOII% = HOY(HOS$)
    fech = FECHATEX$(HOII%)
    TIPODOC$ = "Informe de Produccion: " + NORFA$ + " " + fech ' guardar el documento immpreso
    '
    CONDI$ = "IDSUBOR = '" & NORFA$ & "' and NUMEOPER = " & NUOPE% & " and NUORINF = " & NUMORINF%
    CI% = XVALO(VALOBADA("INFOFAB", "COD_INTE", CONDI$, "NOMESS"))
    FFEE% = CVINT(VALOBADA("INFOFAB", "FEHORLIB", CONDI$, "NOMESS"))
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHATEX$(FFEE%)
    CODPARAM$(2) = "COD-MAT": DOCPARAM$(2) = CODEXT$(CI%)
    CODPARAM$(3) = "DES-MAT": DOCPARAM$(3) = DESCRIT$(CI%)
    CODPARAM$(4) = "NOP-OPER":  DOCPARAM$(4) = VALOBADA("INFOFAB", "NUMEOPER", CONDI$, "NOMESS")
    CODPARAM$(5) = "DES-OPER": DOCPARAM$(5) = VALOBADA("INFOFAB", "DESCOPER", CONDI$, "NOMESS")
    CODPARAM$(6) = "EQUIPO1":  DOCPARAM$(6) = XVALO(VALOBADA("INFOFAB", "MAQUINA", CONDI$, "NOMESS"))
    CODPARAM$(7) = "EQUIPO2":  DOCPARAM$(7) = VALOBADA("INFOFAB", "DESCMAQ", CONDI$, "NOMESS")
    CODPARAM$(8) = "TIE-PROD":  DOCPARAM$(8) = HORATEX$(XVALO(VALOBADA("INFOFAB", "Hora_Has", CONDI$, "NOMESS")))
    CODPARAM$(9) = "TIE-TOTA":  DOCPARAM$(9) = HORMINSE$(VALOBADA("INFOFAB", "TiemEjec", CONDI$, "NOMESS") * 60)
    CODPARAM$(10) = "CANTIDAD": DOCPARAM$(10) = FORMATNUM$(XVALO(VALOBADA("INFOFAB", "CANTREAL", CONDI$, "NOMESS")), "F12.1")
    RETBAJO% = XVALO(VALOBADA("INFOFAB", "RETRABA", CONDI$, "NOMESS"))
    If RETBAJO% = 1 Then
        CODPARAM$(11) = "REF-MAT3": DOCPARAM$(11) = "SI"
        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NUOPE% & ""
        VALOMAX% = XVALO(VALOBADA("INFOFAB", "MAX(NUORINF", CONDI$, "NOMESS"))
        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NUOPE% & " AND NUORINF = " & VALOMAX%
        NMOTV% = XVALO(VALOBADA("INFOFAB", "MOTIRECHA", CONDI$, "NOMESS"))
        CODPARAM$(15) = "REF-MAT4": DOCPARAM$(15) = "Motivo: " & Left$(ECOARCH$("TAMOTIV", Chr(NMOTV%)), 20)
    Else
        CODPARAM$(11) = "REF-MAT3": DOCPARAM$(11) = "NO"
    End If
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CODPARAM$(12) = "REF-MAT2": DOCPARAM$(12) = VALOBADA("ORDEFABR", "ID_ORFA", CONDI$, "NOMESS")
    CODPARAM$(13) = "NUME-COM": DOCPARAM$(13) = NORFA$
    SECCPRO% = XVALO(VALOBADA("ORDEFABR", "SecciPro", CONDI$, "NOMESS"))
    CODPARAM$(14) = "DESTINO": DOCPARAM$(14) = ECOARCH$("SECPROD", MKS$(SECCPRO%))
    CAPARDOC% = 15
    '
    CODTABU1$(1) = "NUOPERS"
    CODTABU1$(2) = "REF-MAT1"
    '
    CACOLTAB1% = 2
    '
    U& = 0
    CAITAB1% = 0
    CONDI$ = "IDSUBOR = '" & NORFA$ & "' and NUMEOPER = " & NUOPE% & " and NUORINF = " & NUMORINF%
    LEGOPER$ = VALOBADA("INFOFAB", "LegaOps", CONDI$, "NOMESS")
    INI% = 1
    For U& = 0 To 3
        OPER% = CVI(Mid$(LEGOPER$, INI%, 2))
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = FORMATNUM$(OPER%, "F3")
        TETABU1$(2, CAITAB1%) = ECOARCH$("PADRON", MKI$(OPER%))
        INI% = INI% + 2
    Next U&
    '
    FORIPRE$ = TRIM$(Control$("", "FORINPR"))
    Call PRINTDOC("@" + FORIPRE$)

    '
End Sub
