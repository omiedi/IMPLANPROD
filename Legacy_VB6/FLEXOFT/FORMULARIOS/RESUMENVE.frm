VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RESUMENVE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Ventas"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10290
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command23 
      Caption         =   "I.Brutos"
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Otros Reportes Mensuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   5250
      TabIndex        =   41
      Top             =   105
      Width           =   4635
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Vendedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   315
         TabIndex        =   70
         Top             =   4830
         Width           =   4005
         Begin VB.PictureBox Picture16 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   71
            Top             =   315
            Width           =   2115
            Begin VB.CommandButton Command21 
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
               Picture         =   "RESUMENVE.frx":103E
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Liquidación     de Comisiones"
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
               TabIndex        =   73
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ventas por Zona / Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   315
         TabIndex        =   43
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   50
            Top             =   1155
            Width           =   2325
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
               Picture         =   "RESUMENVE.frx":1188
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Consulta Totales por Pantalla"
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
               TabIndex        =   52
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   44
            Top             =   735
            Width           =   2325
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
               Picture         =   "RESUMENVE.frx":1492
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Ops. Exentas y Exportación"
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
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   47
            Top             =   315
            Width           =   2325
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
               Left            =   0
               Picture         =   "RESUMENVE.frx":179C
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Operaciones Gravadas"
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
               Index           =   1
               Left            =   735
               TabIndex        =   49
               Top             =   0
               Width           =   1800
            End
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Clasificación por Producto / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   315
         TabIndex        =   42
         Top             =   2415
         Width           =   4005
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   53
            Top             =   1575
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "RESUMENVE.frx":1AA6
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalles por  Cliente y Producto"
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
               Left            =   630
               TabIndex        =   55
               Top             =   105
               Width           =   1620
            End
         End
         Begin VB.PictureBox Picture13 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   59
            Top             =   1155
            Width           =   2325
            Begin VB.CommandButton Command17 
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
               Picture         =   "RESUMENVE.frx":1BF0
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Cliente en Pesos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   735
               TabIndex        =   61
               Top             =   0
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   62
            Top             =   735
            Width           =   2325
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
               Picture         =   "RESUMENVE.frx":2032
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Nro de Factura"
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
               TabIndex        =   64
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   56
            Top             =   315
            Width           =   2325
            Begin VB.CommandButton Command16 
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
               Picture         =   "RESUMENVE.frx":217C
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "Estadisticas de Ventas  y Análisis de Margen Bruto"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísticos"
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
               TabIndex        =   58
               Top             =   0
               Width           =   1410
            End
         End
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   1275
      Left            =   315
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   11
      Top             =   210
      Width           =   4635
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1575
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   1065
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3045
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   750
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   12
         Text            =   " "
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   19
         Top             =   735
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   18
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3885
         TabIndex        =   17
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -735
         TabIndex        =   16
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   13
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Opciones de Salida"
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
      Left            =   5250
      TabIndex        =   7
      Top             =   6405
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   75
         ScaleWidth      =   3945
         TabIndex        =   74
         Top             =   1155
         Width           =   4005
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   735
         Left            =   315
         TabIndex        =   8
         Top             =   315
         Width           =   4005
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
            Left            =   630
            TabIndex        =   10
            Top             =   315
            Value           =   -1  'True
            Width           =   1065
         End
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
            Left            =   2310
            TabIndex        =   9
            Top             =   315
            Width           =   1170
         End
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Proveed"
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":2486
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   7035
      Width           =   855
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":2790
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   5145
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Articulos"
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":2BD2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   6090
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gestión de Cuentas Corrientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   315
      TabIndex        =   3
      Top             =   1575
      Width           =   4635
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   315
         TabIndex        =   65
         Top             =   5040
         Width           =   4005
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3135
            TabIndex        =   66
            Top             =   315
            Width           =   3200
            Begin VB.CommandButton Command22 
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
               Left            =   2580
               Picture         =   "RESUMENVE.frx":2D1C
               Style           =   1  'Graphical
               TabIndex        =   76
               ToolTipText     =   "Re-Imputación de Comprobantes de Ventas"
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
               Left            =   0
               Picture         =   "RESUMENVE.frx":2E66
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command20 
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
               Left            =   2020
               Picture         =   "RESUMENVE.frx":3170
               Style           =   1  'Graphical
               TabIndex        =   67
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Ventas"
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
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Resumen Operativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   315
         TabIndex        =   37
         Top             =   3885
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   38
            Top             =   315
            Width           =   2115
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
               Picture         =   "RESUMENVE.frx":347A
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Men- sual por Cuenta"
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
               TabIndex        =   40
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cobranza y Ajustes"
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
         Left            =   315
         TabIndex        =   30
         Top             =   2310
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   31
            Top             =   735
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
               Left            =   -20
               Picture         =   "RESUMENVE.frx":38BC
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Cobranza en Cuenta Corriente"
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
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   34
            Top             =   315
            Width           =   2535
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
               Picture         =   "RESUMENVE.frx":3BC6
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
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
               Index           =   0
               Left            =   735
               TabIndex        =   36
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   315
         TabIndex        =   20
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   27
            Top             =   1155
            Width           =   2325
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
               Picture         =   "RESUMENVE.frx":3ED0
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación en Cuenta Corriente y de Contado"
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
               Left            =   735
               TabIndex        =   29
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   24
            Top             =   735
            Width           =   2325
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
               Left            =   -20
               Picture         =   "RESUMENVE.frx":41DA
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Débito"
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
               Width           =   1410
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   21
            Top             =   315
            Width           =   2325
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
               Picture         =   "RESUMENVE.frx":44E4
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Crédito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   735
               TabIndex        =   23
               Top             =   105
               Width           =   1695
            End
         End
      End
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":47EE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":4938
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1155
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
      Left            =   10080
      Picture         =   "RESUMENVE.frx":4C42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "RESUMENVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         OPXX% = COMALTER%("Clasificación Según Tipo de Cliente:\\Usuarios\Comerciales\Todos\Cancelar")
         Select Case OPXX%
           Case 1
             Screen.MousePointer = 11
             JJ& = 0
             For U& = 1 To ULTREG&("VEPROVIN")
               X$ = REGLEIDO$("VEPROVIN", U&)
               NCLIE% = CVI(Mid$(X$, 14, 2))
               If CATIBRU$(NCLIE%) <> "3" Then
                 Y$ = String$(64, 0)
                 Call REPLA(Y$, Left$(X$, 13), 1, 21)
                 Call REPLA(Y$, MKI$(NCLIE%), 22, 2)
                 Call REPLA(Y$, CUITCLI$(NCLIE%) + Space$(16), 24, 16)
                 Call REPLA(Y$, Mid$(X$, 16, 8), 40, 8)
                 JJ& = JJ& + 1
                 Call GRAREG("COMULTI", Y$, JJ&)
               End If
             Next U&
             Call SETULTREG("COMULTI", JJ&)
             Call CIERRARCH("COMULTI")
             Call HEADERS("COMULTI", "")
             Call HEADERS("COMULTI", "Clientes: NO COMERCIALES")
             Screen.MousePointer = 1
             Call FINAL("*COMULTI")
             '
           Case 2
             Screen.MousePointer = 11
             JJ& = 0
             For U& = 1 To ULTREG&("VEPROVIN")
               X$ = REGLEIDO$("VEPROVIN", U&)
               NCLIE% = CVI(Mid$(X$, 14, 2))
               If CATIBRU$(NCLIE%) = "3" Then
                 Y$ = String$(64, 0)
                 Call REPLA(Y$, Left$(X$, 13), 1, 21)
                 Call REPLA(Y$, MKI$(NCLIE%), 22, 2)
                 Call REPLA(Y$, CUITCLI$(NCLIE%) + Space$(16), 24, 16)
                 Call REPLA(Y$, Mid$(X$, 16, 8), 40, 8)
                 JJ& = JJ& + 1
                 Call GRAREG("COMULTI", Y$, JJ&)
               End If
             Next U&
             Call SETULTREG("COMULTI", JJ&)
             Call CIERRARCH("COMULTI")
             Call HEADERS("COMULTI", "")
             Call HEADERS("COMULTI", "Clientes: COMERCIALES")
             Screen.MousePointer = 1
             Call FINAL("*COMULTI")
             '
           Case 3
             Call FINAL("*VEPROLI")
         End Select
    End If
End Sub

Private Sub Command11_Click()
    Call OPNOIN("")
End Sub

Private Sub Command12_Click()
    '
    If Option1.Value = True Then
         MOX% = ALTERNA%("Consulta Selectiva por Cliente\Consulta por Codigo\Consulta General por Cliente\Consulta General por Pieza")
         On MOX% GoTo 203, 204, 205, 206
       ElseIf Option2.Value = True Then
         MOX% = ALTERNA%("Listado General por Cliente\Listado por Codigo de Producto")
         On MOX% GoTo 201, 202
    End If
    Exit Sub
    '
201 Sali$ = "*REVECLI": GoTo 999
202 Sali$ = "*REVEPIE": GoTo 999
203 Sali$ = "*COVECLI": GoTo 999
204 Sali$ = "*COVEPIE": GoTo 999
205 Sali$ = "*COGECLI": GoTo 999
206 Sali$ = "*COGEPIE": GoTo 999
    '
999 Call FINAL(Sali$)
    '
End Sub

Private Sub Command16_Click()
    MOX% = ALTERNA%("Resumen Estadistico de Ventas\Margen Bruto de Ventas")
    On MOX% GoTo 203, 204
    Exit Sub
    '
203 Call CONECRUN("#GESTAVE", "Resumen Estadistico de Ventas")
    Exit Sub
    '
204 Call CONECRUN("#MABRUVE", "Analisis Margen Bruto de Ventas")
    Exit Sub
    '
End Sub

Private Sub Command17_Click()
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call FINAL("*RESUSOR")
    End If
End Sub

Private Sub Command18_Click()
    Call CONECRUN("#GEDEFAC", "Detalle por Número de Factura")
End Sub

Private Sub Command19_Click()
    If Option1.Value = True Then
         Call CONECRUN("*COSUDIV", "Consulta por Cuenta Contable")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call CONECRUN("*SOSUDIV", "Listado de Ventas por Cuenta Contable")
         Exit Sub
    End If
End Sub

Private Sub Command2_Click()
    MOX% = ALTERNA%("Secuencia por Recibo\Detalle por Cliente")
    On MOX% GoTo 203, 204
    Exit Sub
    '
203 Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETCOBV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(RESUMENVE.MESAN.TEXT)
    C2$ = TRIM$(RESUMENVE.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENVE.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
    Exit Sub
    '
204 Call HEADERS("RECOBCLI", "")
    Call HEADERS("RECOBCLI", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("RECOBCLI", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*COCOBCL")
      ElseIf Option2.Value = True Then
         Call FINAL("*SOCOBCL")
    End If
    Exit Sub
    '
End Sub
'

Private Sub Command20_Click()
    Call CONASIDIA
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    ATTX% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MAINIERI") > 0 Then
        Call CONECRUN("LIQCOMAI", "Liquidacion de Comisiones")
        GoTo 99
      ElseIf InStr(EPAC$, "SEALU") > 0 Then
        Call CONECRUN("LIQCOMAI", "Liquidacion de Comisiones")
        GoTo 99
      ElseIf InStr(EPAC$, "DOSIVAC") > 0 Then
        Call CONECRUN("LICODOSI", "Liquidacion de Comisiones")
        GoTo 99
    End If
    '
    FORLICOMIS.MESAN.TEXT = MESAN.TEXT
    FORLICOMIS.Text4(0).TEXT = Text4(0).TEXT
    FORLICOMIS.Text4(1).TEXT = Text4(1).TEXT
    FORLICOMIS.Show 1
99  Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   '
   Command22.Enabled = False
   '
5  CORRESP1$ = TRIM$(MESAN.TEXT)
   CORRESQ1$ = MESANO$(CORRESP1$, DES%, HAS%)
   FEX = DES%: DESFE1$ = FECHATEX$(FEX)
   FEX = HAS%: HASFE1$ = FECHATEX$(FEX)
   '
   CORRESQ1$ = TRIM$(CORRESP1$)
   ABREMES1$ = Left$(CORRESQ1$, 3) + Mid$(DESFE1$, 7, 2)
   LUXI$ = LUDAT$
   Archi$ = "ICV" + ABREMES1$ + ".DAT"
   N2% = FILEOPEN(LU$ + Archi$, 0, 96)
   Dim BULEC As String * 96
   '
10 NUFACLI$ = REGSEL$("IFACVEN")
   If Len(NUFACLI$) <= 4 Then
      Close #N2%
      GoTo 99
   End If
   NF$ = Left$(NUFACLI$, 18)
   NC% = CVI(Mid$(NUFACLI$, 59, 2))
   IMPO# = CVD(Mid$(NUFACLI$, 51, 8))
   NRO& = CVS(Mid$(NUFACLI$, 61, 4))
   If Abs(IMPO#) < 0.005 Then
      Call MENSERR(24, "Documento Anulado")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   For U& = 1 To LOF(N2%) / 96
     Get #N2%, U&, BULEC$
     If Left$(BULEC$, 2) = Left$(NF$, 2) Then
       If CVS(Mid$(BULEC$, 3, 4)) = NRO& Then
         If CVI(Mid$(BULEC$, 7, 2)) = NC% Then
           REG& = U&
           Y6$ = Mid$(BULEC$, 17, 80)
           GoTo 20
         End If
       End If
     End If
   Next U&
   Close #N2%
   Call MENSERR(24, "Documento no Encontrado - Consulte.")
   GoTo 99
   '
20 ATRI$ = TRIM$(Left$(NUFACLI$, 50)) + " ($ " + TRIM$(CSTRING$(MKD$(IMPO#), 4, 12, 2, 2)) + ")"
   '
   Call BLANARCH("!CARICOM")
   For J% = 1 To 8
      IPA# = 0
      CUE% = CVI(Mid$(Y6$, 10 * J% - 9, 2))
      If CUE% > 0 Then
         IPA# = CVD(Mid$(Y6$, 10 * J% - 7, 8))
         Y$ = REGBLAN$("!CARICOM")
         Call REPLA(Y$, MKI$(CUE%), 1, 2)
         Call REPLA(Y$, MKD$(IPA#), 3, 8)
         Call REGAPP("!CARICOM", Y$)
      End If
   Next J%
   Call CIERRARCH("!CARICOM")
   '
30 VTB% = VENTABU%("CARIFAVE/" + ATRI$)
   If VTB% < 0 Then GoTo 10
   '
   TOTIMPU# = 0: Y6$ = ""
   For U& = 1 To ULTREG&("!CARICOM")
       Y$ = REGLEIDO$("!CARICOM", U&)
       CUEB$ = Left$(Y$, 2)
       If CVI(CUEB$) > 0 Then
          DECUE$ = ECOARCH$("CUECON", CUEB$)
          If DECUE$ <> "" Then
             IMBB$ = Mid$(Y$, 3, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 30
             End If
             Y6$ = Y6$ + CUEB$ + IMBB$
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
          End If
       End If
   Next U&
   '
   If Abs(Abs(IMPO#) - TOTIMPU#) >= 0.005 Then
      Call MENSERR(24, "Suma de Imputaciones Difiere de Total")
      GoTo 30
   End If
   '
   Y6$ = Y6$ + String$(80, 0)
   Get #N2%, REG&, BULEC$
   X$ = BULEC$
   Call REPLA(X$, Y6$, 17, 80)
   BULEC$ = X$
   Put #N2%, REG&, BULEC$
   Close #N2%
   GoTo 5
   '
99 Call FRESHECHO("IFACVEN")
   Call COMUNI("Para Efectivizar los Cambios\Debe Recalcular el Resumen Mensual Ahora.")
   Command22.Enabled = True
   Hide
   '
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
   If OPX% = 1 Then
       Call HEADERS("SUPERIB", "")
       Call HEADERS("SUPERIB", "Corresponde a: " + MESAN)
       Call FINAL("*SUPERIB")
     ElseIf OPX% = 2 Then
       CommonDialog1.CancelError = True
       CommonDialog1.DialogTitle = "Crear Archivo de Transferencia"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
       CommonDialog1.FILTER = "Todos los Archivos (*.*)|*.*|"
       CommonDialog1.FilterIndex = 1
       On Error GoTo 99
       CommonDialog1.ShowOpen
       On Error GoTo 0
       '
       ASALID$ = UCase$(CommonDialog1.filename)
       NX% = FILEOPEN%(ASALID$, 2, 48)
       CAREGI% = 0
       For U& = 1 To ULTREG&("SUPERIB")
         XX$ = REGLEIDO$("SUPERIB", U&)
         TTXX$ = Space$(60)
           CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
         Call REPLA(TTXX$, CUITTX$, 1, 13)
           FEX = CVI(Left$(XX$, 2))
           FEXU$ = FECHATEX$(FEX)
           FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         Call REPLA(TTXX$, FEXU$, 14, 10)
           LECOX$ = "F"
           If Mid$(XX$, 3, 2) = "ND" Then LECOX$ = "D"
           If Mid$(XX$, 3, 2) = "NC" Then LECOX$ = "C"
         Call REPLA(TTXX$, LECOX$, 24, 1)
           LECO$ = Mid$(XX$, 20, 1)
         Call REPLA(TTXX$, LECO$, 25, 1)
           SUCUX$ = Mid$(XX$, 6, 4)
         Call REPLA(TTXX$, SUCUX$, 26, 4)
           NUCOX$ = Mid$(XX$, 11, 8)
         Call REPLA(TTXX$, NUCOX$, 30, 8)
           IMPOL# = CVD(Mid$(XX$, 79, 8))
           IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
           While Len(IMPOX$) < 12: IMPOX$ = "0" + IMPOX$: Wend
           If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
         Call REPLA(TTXX$, IMPOX$, 38, 12) 'Esto se comenta por mificacion en programa ARIB(AFIP)06/05/09
           IMPOL# = CVD(Mid$(XX$, 87, 8))
           IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
           While Len(IMPOX$) < 11: IMPOX$ = "0" + IMPOX$: Wend
           If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
         Call REPLA(TTXX$, IMPOX$, 50, 11)
         Call REPLA(TTXX$, "A", 61, 1)
         'Call REPLA(TTXX$, IMPOX$, 38, 11) 'Esto se modifica por mificacion en programa ARIB(AFIP)06/05/09
         Print #NX%, TTXX$
         CAREGI% = CAREGI% + 1
       Next U&
       Close #NX%
       ASALID1$ = REPLACAR$(ASALID$, "\", "/")
       Call COMUNI("Archivo de Transferencia '" + ASALID$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
   End If
99 Command23.Enabled = True
   On Error GoTo 0
End Sub

Private Sub Command3_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETFACV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(RESUMENVE.MESAN.TEXT)
    C2$ = TRIM$(RESUMENVE.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENVE.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Private Sub Command4_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETDEBV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(RESUMENVE.MESAN.TEXT)
    C2$ = TRIM$(RESUMENVE.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENVE.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Notas de Debito")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub


Private Sub Command5_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETCREV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(RESUMENVE.MESAN.TEXT)
    C2$ = TRIM$(RESUMENVE.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENVE.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Notas de Credito")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Private Sub Command6_Click()
    If Option1.Value = True Then
         Call FINAL("*RESUCCD")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCD")
    End If
End Sub

Private Sub Command7_Click()
    Screen.MousePointer = 11
    LU$ = LUDAT$
    FileCopy LU$ + "DETOTRV.DAT", LU$ + "DETACCX.DAT"
    C1$ = TRIM$(RESUMENVE.MESAN.TEXT)
    C2$ = TRIM$(RESUMENVE.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENVE.Text4(1).TEXT)
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Otros Movimientos de Cuenta Corriente")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGEVEN")
End Sub

Private Sub Command9_Click()
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call FINAL("*VEPROLX")
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    '\\\OT-06-0559-RG-19/12/06///
    If InStr(UCase$(EPAC$), "SABO") Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(UCase$(EPAC$), "FERVI") > 0 Then
    '\\\OT-06-0559-RG-FIN///
      Command23.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub MESAN_GOTFOCUS()
    Command1.SetFocus
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command1.SetFocus
End Sub
