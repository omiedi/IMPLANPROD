VERSION 5.00
Begin VB.Form COGEVENT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Gestión Comercial - Cuentas por Cobrar"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10455
   FillStyle       =   0  'Solid
   Icon            =   "COGEVENT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Listados de Control Estadistico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2050
      Left            =   4830
      TabIndex        =   49
      Top             =   210
      Width           =   4425
      Begin VB.CommandButton Command21 
         Caption         =   "R.D.M.P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3255
         TabIndex        =   63
         ToolTipText     =   "Resumen Diario Movimiento de Piezas (SABO)"
         Top             =   420
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   50
         Top             =   1260
         Width           =   2325
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
            Picture         =   "COGEVENT.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Detalle de ventas por Cliente - Producto - Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
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
            TabIndex        =   52
            Top             =   105
            Width           =   1620
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   53
         Top             =   840
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
            Left            =   -20
            Picture         =   "COGEVENT.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Detalle de Ventas en Pesos por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
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
            TabIndex        =   55
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
         TabIndex        =   56
         Top             =   420
         Width           =   2325
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
            Picture         =   "COGEVENT.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Estadisticas de Ventas y Análisis de Margen Bruto"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
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
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9420
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   60
      Top             =   2130
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4620
      Sorted          =   -1  'True
      TabIndex        =   59
      Top             =   1365
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command19 
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command18 
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":0CEA
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command16 
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":112C
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Alta de Nuevo Proveedor"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0FF&
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
      Height          =   4320
      Left            =   4830
      TabIndex        =   6
      Top             =   2415
      Width           =   4425
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Configuración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   24
            Top             =   795
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
               Left            =   0
               Picture         =   "COGEVENT.frx":1436
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Tablas de validación de Datos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Tablas de Validación"
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
               TabIndex        =   25
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   21
            Top             =   375
            Width           =   2325
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
               Left            =   -20
               Picture         =   "COGEVENT.frx":1740
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Diseñp de Formularios de Impresión"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Diseño de Formularios"
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
               TabIndex        =   23
               Top             =   0
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D0F0FF&
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
         Height          =   2010
         Left            =   420
         TabIndex        =   7
         Top             =   2100
         Width           =   3585
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   35
            Top             =   1225
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
               Left            =   -20
               Picture         =   "COGEVENT.frx":1A4A
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes y Deudores "
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
               TabIndex        =   37
               Top             =   105
               Width           =   1065
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   32
            Top             =   805
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
               Picture         =   "COGEVENT.frx":1E8C
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Artículos de Stock"
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
               TabIndex        =   34
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   29
            Top             =   385
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
               Left            =   -20
               Picture         =   "COGEVENT.frx":2196
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuentas Contables"
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
               TabIndex        =   31
               Top             =   0
               Width           =   960
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Ventas y Deudores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Resúmenes Mensuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   420
         TabIndex        =   42
         Top             =   5145
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   43
            Top             =   370
            Width           =   2325
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
               Picture         =   "COGEVENT.frx":24A0
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Generación y Emisión de Listados y Estadísticas Mensuales."
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Subdiarios y Estadística"
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
               TabIndex        =   45
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Control de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   420
         TabIndex        =   5
         Top             =   1680
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   46
            Top             =   1680
            Width           =   2115
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
               Left            =   -20
               Picture         =   "COGEVENT.frx":27AA
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Listados de Avance de Facturación y Cobranza por Día"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Avance Factura- cion y Cobranza"
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
               TabIndex        =   48
               Top             =   105
               Width           =   1500
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Salida"
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
            Left            =   210
            TabIndex        =   26
            Top             =   2460
            Width           =   3165
            Begin VB.OptionButton Option2 
               BackColor       =   &H00D0F0FF&
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
               Left            =   1785
               TabIndex        =   28
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00D0F0FF&
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
               Left            =   315
               TabIndex        =   27
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   18
            Top             =   1260
            Width           =   2220
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
               Picture         =   "COGEVENT.frx":2AB4
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Listado de Saldos y Deudas de Clientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldos y Deudas de Clientes"
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
               TabIndex        =   20
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   15
            Top             =   840
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
               Left            =   0
               Picture         =   "COGEVENT.frx":2DBE
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Listado de Planificación Financiera de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Plan Financiero"
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
               TabIndex        =   17
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   12
            Top             =   420
            Width           =   2430
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
               Picture         =   "COGEVENT.frx":30C8
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Listados de Facturas Pendientes de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuentas por Cobrar"
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
               TabIndex        =   14
               Top             =   105
               Width           =   1905
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Detalle por Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   9
            Top             =   370
            Width           =   2955
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
               Left            =   2340
               Picture         =   "COGEVENT.frx":33D2
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Listado de Límites de Crédito y Deudas"
               Top             =   0
               Width           =   560
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
               Left            =   0
               Picture         =   "COGEVENT.frx":36DC
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Cliente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes de Cuenta Corriente"
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
               TabIndex        =   11
               Top             =   105
               Width           =   1485
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":39E6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   340
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":3B30
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1180
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
      Height          =   800
      Left            =   9420
      Picture         =   "COGEVENT.frx":3E3A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8880
      Picture         =   "COGEVENT.frx":4144
      Top             =   6210
      Width           =   2010
   End
End
Attribute VB_Name = "COGEVENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    ACONEC$ = "AMASTO01"
    If EXISTE%("AMASTO02.EXE") = 1 Then
      ACONEC$ = "AMASTO02"
    End If
    Call CONECRUN(ACONEC$, "Maestro de Items de Stock")
    Command11.Enabled = True
End Sub


Private Sub Command12_Click()
    Command12.Enabled = False
    ACONEC$ = "AMACLI01"
    If EXISTE%("AMACLI02.EXE") = 1 Then
      ACONEC$ = "AMACLI02"
    End If
    Call CONECRUN(ACONEC$, "Actualización de Base de Datos de Clientes")
    Command12.Enabled = True
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call MENSERR(24, "Imposible Acceder desde esta Pantalla.\  \Pase al Escritorio de\Actualización de Archivos Maestros.")
    Command10.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
       ElseIf Option2.Value = True Then
         HOO% = HOY(HO$)
         VDEF$ = MKI$(HOO%) + MKI$(HOO%)
         OPX$ = OBJPLA$("DESHASFECHA", VDEF$)
         If OPX$ = "" Then Exit Sub
         DES% = CVI(Left$(OPX$, 2))
         HAS% = CVI(Mid$(OPX$, 3, 2))
         If DES% > HAS% Then Exit Sub
         Call CALREVECLI(DES%, HAS%)
    End If
    '
End Sub
'

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "SABO") > 0 Then GoTo 10
    '
    If EXISTE%("IGESVE02.EXE") > 0 Then
         Call CONECRUN("IGESVE02", "Informe de Gestion de Ventas")
       Else
10       RESUESTA.Show 1
    End If
    Command15.Enabled = True
    '
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    HOO% = HOY(HO$)
    VDEF$ = MKI$(HOO%) + MKI$(HOO%)
    OPX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OPX$ = "" Then GoTo 99
    DES% = CVI(Left$(OPX$, 2))
    HAS% = CVI(Mid$(OPX$, 3, 2))
    If DES% > HAS% Then GoTo 99
    OSCREEN% = 0: If Option1.Value = True Then OSCREEN% = 1
    Call CALRESUVE(DES%, HAS%)
    '
99  Command14.Enabled = True
    '
End Sub

Private Sub Command16_Click()
    Call CONECRUN("AMAPRO", "Maestro de Proveedores")
    Call FRESHECHO("PROVED1")
End Sub

Private Sub Command17_Click()
    Call OPNOIN("")
End Sub

Private Sub Command18_Click()
   OPCLAVEN.Show 1
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    ACONEC$ = "AMASTO01"
    If EXISTE%("AMASTO02.EXE") = 1 Then
      ACONEC$ = "AMASTO02"
    End If
    Call CONECRUN(ACONEC$, "Maestro de Items de Stock")
    Command19.Enabled = True
End Sub

'
Private Sub Command2_Click()
    CARGAMEN% = 0
    REMEVENT.Show 1
End Sub

Private Sub Command20_Click()
   Command20.Enabled = False
   Call LILIMCRE
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   'OK MOVISVIE
   Command21.Enabled = False
   Dim AREMO$(3), STMOVX$(3), JK1&(3), TOENTRA#(3), TOSALE#(3), LEDES$(3)
   '
   JK1&(1) = 0: JK1&(2) = 0: JK1&(3) = 0
   AREMO$(1) = "REDIMOV": STMOVX$(1) = "FC-RT-RD-RM-ND-DV"
   AREMO$(2) = "REDIMOR": STMOVX$(2) = "BR"
   AREMO$(3) = "REDIMOF": STMOVX$(3) = "OF-PF-CF-CP"
   '
   LEDES$(1) = "Ventas - Facturación"
   LEDES$(2) = "Boletas de Recepción"
   LEDES$(3) = "Fabricación / Armado Kits"
   For KKI% = 1 To 3
      Call COPYSTRU("REDIMOST", AREMO$(KKI%))
      TOENTRA#(KKI%) = 0: TOSALE#(KKI%) = 0
   Next KKI%
   '
   Call SELECHO("SELFECHA")
   FERESU$ = VALACT1$("SELFECHA")
   If FERESU$ <> "" Then
     FF% = FECHANUM(FERESU$)
     If FF% > 0 Then
       ARMOVI$ = "MOVISTO"
       '
13     FIN& = ULTREG&(ARMOVI$)
       '
       Screen.MousePointer = 11
       REBLA$ = REGBLAN$("REDIMOST")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$(ARMOVI$, U&)
         FEMI% = CVI(Left$(XX$, 2))
         If FEMI% = FF% Then
           CIXI% = CVI(Mid(XX$, 3, 2))
           If CIXI% > 0 Then
             TIMO$ = UCase$(TRIM$(Mid$(XX$, 21, 2)))
             If TIMO$ <> "" Then
               For KKI% = 1 To 3
                 If InStr(STMOVX$(KKI%), TIMO$) > 0 Then
                   CAENTRA# = CVD(Mid$(XX$, 88, 8))
                   CASALE# = CVD(Mid$(XX$, 96, 8))
                   TTXX$ = REBLA$
                   Call REPLA(TTXX$, Mid$(XX$, 23, 50), 1, 50)
                   Call REPLA(TTXX$, MKI$(CIXI%), 51, 2)
                   Call REPLA(TTXX$, MKD$(CAENTRA#), 53, 8)
                   TOENTRA#(KKI%) = TOENTRA#(KKI%) + CAENTRA#
                   Call REPLA(TTXX$, MKD$(CASALE#), 61, 8)
                   TOSALE#(KKI%) = TOSALE#(KKI%) + CASALE#
                   JK1&(KKI%) = JK1&(KKI%) + 1
                   Call GRAREG(AREMO$(KKI%), TTXX$, JK1&(KKI%))
                   GoTo 19
                 End If
               Next KKI%
             End If
           End If
         End If
19     Next U&
       '
       If ARMOVI$ = "MOVISTO" Then
          If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
            XXX$ = REGLEIDO$("MOVISTO", 1)
            FF11% = CVI(Left$(XXX$, 2))
            If FF11% >= FF% Then
              ARMOVI$ = "MOVISVIE"
              GoTo 13
            End If
          End If
       End If
       
       For KKI% = 1 To 3
         Call SETULTREG(AREMO$(KKI%), JK1&(KKI%))
         Call CIERRARCH(AREMO$(KKI%))
       Next KKI%
       '
       JJ& = 0
       For KU% = 1 To 3
         XX$ = REGBLAN$("REDMOST")
         Call REPLA(XX$, Left$(STMOVX$(KU%), 2), 1, 2)
         Call REPLA(XX$, LEDES$(KU%), 3, 30)
         Call REPLA(XX$, MKD$(TOENTRA#(KU%)), 33, 8)
         Call REPLA(XX$, MKD$(TOSALE#(KU%)), 41, 8)
         Call REPLA(XX$, MKD$(TOENTRA#(KU%) - TOSALE#(KU%)), 49, 8)
         JJ& = JJ& + 1
         Call GRAREG("REDMOST", XX$, JJ&)
       Next KU%
       Call SETULTREG("REDMOST", JJ&)
       Call HEADERS("REDMOST", "")
       Call HEADERS("REDMOST", "Corresponde a: " + FERESU$)
       For KKI% = 1 To 3
         Call HEADERS(AREMO$(KKI%), "")
         Call HEADERS(AREMO$(KKI%), "Corresponde a: " + FERESU$)
       Next KKI%
       Call FINAL("*REDMOST")
       '
     End If
   End If
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
For U& = 1 To ULTREG&("MENSAJES")
XX$ = REGLEIDO$("MENSAJES", U&)
If InStr(XX$, "5018") > 0 Then
FEX = CVI(Mid$(XX$, 5, 2))
MsgBox FECHATEX$(FEX) + "    " + Mid$(XX$, 13, 30) + Mid$(XX$, 75, 82)
AAA = BBB
End If
Next U&
MsgBox "FIN"
End Sub

Private Sub Command3_Click()
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32767), PUNSALDEU$(32767)
    '
10  RECONMOVI& = 0
    '
    If RECLEN%("FAPECOB") < 64 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta\Archivo 'FAPECOB' - Pasar a 64 by.")
       Screen.MousePointer = 1
       Exit Sub
    End If
    '
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("SVD202", U&)
      TIX% = Asc(Mid$(X$, 3, 1))
      If TIX% = 7 Then TIX% = 0
      IM1# = CVD(Mid$(X$, 11, 8))
      IM2# = CVD(Mid$(X$, 21, 8))
      If TIX% > 0 Then
        If IM1# <> Abs(IM1#) Or IM2# <> Abs(IM2#) Then
          IM1# = Abs(IM1#): IM2# = Abs(IM2#)
          Call REPLA(X$, Chr$(TIX%), 3, 1)
          Call REPLA(X$, MKD$(IM1#), 11, 8)
          Call REPLA(X$, MKD$(IM2#), 21, 8)
          Call GRAREG("SVD202", X$, U&)
        End If
      End If
      '
      MONEMX% = Asc(Mid$(X$, 31, 1))
      If MONEMX% < 1 Or MONEMX% > 2 Then GoSub 200
      '
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        If NC% < 32768 Then
          If Abs(IM1#) >= 0.005 Then
            SIG% = 1: If TIX% > 3 And TIX% < 7 Then SIG% = (-1)
            SALDEUPEN#(NC%) = SALDEUPEN#(NC%) + SIG% * IM1#
            PUNSALDEU$(NC%) = PUNSALDEU$(NC%) + MKS$(U&)
            RECONMOVI& = RECONMOVI& + 1
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("SVD202")
    '
    URENUE& = RECONMOVI&
    UREVIE& = ULTREG&("SVD202")
    '
    If URENUE& < UREVIE& * 0.8 Then
       TTXX$ = "Hay Mucho Espacio Desaprovechado\en la Base de Datos de Facturas Pendientes de Cobranza."
       TTXX$ = TTXX$ + "\  \Compactar esta Base de Datos Puede Acelerar\Significativamente el Funcionamiento del Sistema."
       TTXX$ = TTXX$ + "\  \Para Efectuar la Compactación, TODOS los Restantes Usuarios\Deben Salir de FLEXOFT hasta el Escritorio Principal."
       TTXX$ = TTXX$ + "\  \Desea Ahora Compactar la Base de Datos\de Facturas Pendientes de Cobranza ?"
       If COMALTER%(TTXX$ + "\\Continuar\Compactar Ahora") = 2 Then
          Call COMPACDEU
          For KKIL& = 0 To 32767
            SALDEUPEN#(KKIL&) = 0
            PUNSALDEU$(KKIL&) = ""
          Next KKIL&
          GoTo 10
       End If
    End If
    '
    FIN& = ULTREG&("DEUDOR1")
    Screen.MousePointer = 11
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DEUDOR1", U&)
      NC% = CVI(Left$(X$, 2))
      If NC% > 0 Then
        If NC% <= 32767 Then
          X2$ = REGLEIDO$("DEUDOR2", U&)
          SCLX# = CVD(Mid$(X2$, 109, 8))
          If Abs(SCLX# - SALDEUPEN#(NC%)) >= 0.005 Then
            GoSub 100
          End If
        End If
      End If
    Next U&
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPECOBRA.Show 1
    Exit Sub
    '
100 SALDEUPEN#(NC%) = 0
    HABIAREPE% = 0
    List1.Clear
    XXXX = Len(PUNSALDEU$(NC%))
    For U1& = 1 To Len(PUNSALDEU$(NC%)) Step 4
       REG1& = CVS(Mid$(PUNSALDEU$(NC%), U1&, 4))
       X1$ = REGLEIDO$("SVD202", REG1&)
       SUBIT% = 0
       For KKI% = 1 To Len(X1$)
          SUBIT% = SUBIT% + Asc(Mid$(X1$, KKI%, 1))
       Next KKI%
       TTXX$ = Space$(20)
       Call REPLA(TTXX$, Str$(SUBIT%), 1, 10)
       Call REPLA(TTXX$, Str$(U1&), 11, 10)
       Call REPLA(TTXX$, "A", 1, 1)
       List1.AddItem TTXX$
    Next U1&
    '
    For U1& = 1 To List1.ListCount
       Y1$ = List1.List(U1& - 1)
       For U2& = U1& + 1 To List1.ListCount
          Y2$ = List1.List(U2& - 1)
          If Left$(Y1$, 10) <> Left$(Y2$, 10) Then GoTo 110
          REG1& = Val(Mid$(Y1$, 11))
          REG2& = Val(Mid$(Y2$, 11))
          X1$ = REGLEIDO$("SVD202", REG1&)
          X2$ = REGLEIDO$("SVD202", REG2&)
          If X1$ = X2$ Then
            Call REPLA(X2$, String$(30, 0), 3, 30)
            Call GRAREG("SVD202", X2$, REG2&)
            HABIAREPE% = 1
          End If
       Next U2&
110 Next U1&
Return
    
200 NC% = CVI(Left$(X$, 2))
    IMPO# = CVD(Mid$(X$, 21, 8))
    SALDO# = CVD(Mid$(X$, 11, 8))
    NROL& = CVS(Mid$(X$, 5, 4))
    MONEMX% = Asc(Mid$(X$, 31, 1))
    '
    If MONEMX% > 2 Then
      MONEMX% = 0
      Call REPLA(X$, Chr$(0), 31, 1)
      Call GRAREG("SVD202", X$, U&)
    End If
    '
    If MONEMX% <= 1 Then
      RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
      For KUL& = Len(RFG$) - 3 To 1 Step -4
        RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
        XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
        If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
          If Asc(Mid$(XXC$, 37, 1)) = TIX% Then
            MONEMX% = Asc(Mid$(XXC$, 31, 1))
            GoTo 215
          End If
        End If
      Next KUL&
    End If
    '
215 If MONEMX% <> 2 Then MONEMX% = 1
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, U&)
    '
Return
'
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If EXISTE%("INFOPAT.EXE") > 0 Then
          Call CONECRUN("INFOPAT", "Informe Patrimonial Diario")
        Else
          Call MENSERR(24, "Opción no Disponible\en la Presente Versión FLEXOFT.\  \Utilice las Opciones Habilitadas\en el Módulo de Cuentas por Cobrar.")
          'Call CONECRUN("#SVD47", "Plan Financiero de Cobranza")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALCLI
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CONECRUN("PRIFORMS", "Diseño de Formularios de Impresion")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    'Hide
    RESCUDEU.Show
    RESCUDEU.Refresh
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGEVEN")
End Sub

Private Sub Command9_Click()
    FORSELME.Show 1
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
    '
    If InStr(UCase$(EPAC$), "SABO") > 0 Then
      Command21.Visible = True
    End If
    '
    'SVLC$ = SERIVOL$("C:")
    If InStr(UCase$(EMPREAC$), "NOVOESPU") > 0 Or InStr(UCase$(EMPREAC$), "DOSIVAC") > 0 Or SVLC$ = "081D-0BD9" Or SVLC$ = "2560-10F0" Then
        Command4.ToolTipText = "Informe Patrimonial Diario"
        Label5.Caption = "Inf. Patrimonial"
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Sub COMPACDEU()
   '
   Call CIERRARCH("*.*")
   Call CIERRARCH("!SVD202")
   Call BLANARCH("!SVD202")
   JJ& = 0
   '
   FIN& = ULTREG&("SVD202")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("SVD202", U&)
     IM1# = CVD(Mid$(X$, 11, 8))
     NC% = CVI(Left$(X$, 2))
     If NC% > 0 Then
       If NC% <= 32767 Then
         If Abs(IM1#) >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("!SVD202", X$, JJ&)
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("!SVD202")
   '
   On Error GoTo 10
   Kill LUDAT$ + "SVD202.X*"
   Kill LUDAT$ + "SVD202.DAT"
   On Error GoTo 0
   '
   FIN& = ULTREG&("!SVD202")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!SVD202", U&)
     JJ& = JJ& + 1
     Call GRAREG("SVD202", X$, JJ&)
   Next U&
   Call CIERRARCH("SVD202")
   Call CIERRARCH("*.*")
   Call COMUNI("Compactación Completa.")
   Exit Sub
   '
10 Resume 20
20 On Error GoTo 0
   Call COMUNI("La Compactación no Pudo Completarse.")
   '
End Sub

Sub CALRESUVE(DES%, HAS%)
    '
    Dim TIS$(10)
    '
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    Next I%
    FIN& = ULTREG&("FACTUR")
    '
    If Option1.Value = True Then ' por pantalla
10     RDEU$ = REGSEL$("DEUDOR1")
       If Len(RDEU$) <= 4 Then Exit Sub
       NC% = CVI(Left$(RDEU$, 2))
       RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
       If Len(RFF$) < 4 Then
15       Call MENSERR(24, "Sin Movimiento")
         GoTo 10
       End If
       '
       CLI$ = AJUSTI$(TRIM$(Str$(NC%)) + " - " + RASOCLI$(NC%), 30)
       Call COPYSTRU("RESUVEN", "CRESVEN")
       REBLA$ = REGBLAN$("RESUVEN")
       JJ& = 0
       '
       For U& = 1 To Len(RFF$) Step 4
         IL& = CVS(Mid$(RFF$, U&, 4))
         X$ = REGLEIDO$("FACTUR", IL&)
         FF% = CVI(Mid$(X$, 7, 2))
         '
         If FF% >= DES% And FF% <= HAS% Then
           TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
           VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
           TIXX$ = TIS$(TI%)
           NRO& = CVS(Mid$(X$, 3, 4))
           '
           NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100
           '
           SIG% = 1
           If TI% > 3 Then
             NETO# = Abs(NETO#)
             SIG% = (-1)
           End If
           '
           If TI% > 0 Then
             If TI% <= 6 Then
               YR$ = REBLA$
               JI% = JJ& + 1
               Call REPLA$(YR$, MKI$(JI%), 1, 2)
               Call REPLA$(YR$, CLI$, 3, 30)
               Call REPLA(YR$, TIXX$ + AJUSTD$(Str$(NRO&), 6), 33, 8)
               Call REPLA(YR$, MKI$(FF%), 41, 2)
               Call REPLA(YR$, MKD$(NETO# * SIG%), 43, 8)
               If SIG% < 0 Then
                 Call REPLA(YR$, "Credito", 59, 8)
               End If
               JJ& = JJ& + 1
               Call GRAREG("CRESVEN", YR$, JJ&)
             End If
           End If
         End If
         '
       Next U&
       '
       Call SETULTREG&("CRESVEN", JJ&)
       Call CIERRARCH("CRESVEN")
       '
       If JJ& < 1 Then GoTo 15 ' sin movimiento
       Call FINAL("*CRESVEN/" + CLI$)
       GoTo 10
       '
    End If
    '
    Screen.MousePointer = 11
    CARESUVE& = 0
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("FACTUR", IL&)
        NC% = CVI(Left$(X$, 2))
        FF% = CVI(Mid$(X$, 7, 2))
        '
        If FF% >= DES% And FF% <= HAS% Then
           TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
           VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
           TIXX$ = TIS$(TI%)
           NRO& = CVS(Mid$(X$, 3, 4))
           '
           NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100
           '
           SIG% = 1
           If TI% > 3 Then
               NETO# = Abs(NETO#)
               SIG% = (-1)
           End If
           '
           If TI% > 0 Then
               If TI% <= 6 Then
                   If NC% >= 0 Then
                         CLI$ = RASOCLI$(NC%)
                       ElseIf NC% < 0 Then
                         RECUI& = Abs(NC%)
                         If RECUI& <= ULTREG&("CUITS") Then
                             RCUX$ = REGLEIDO$("CUITS", RECUI&)
                             CLI$ = Left$(RCUX$, 30)
                         End If
                   End If
                   '
                   YR$ = REGBLAN$("RESUVEN")
                   Call REPLA$(YR$, MKI$(NC%), 1, 2)
                   Call REPLA$(YR$, CLI$, 3, 30)
                   Call REPLA(YR$, TIXX$ + AJUSTD$(Str$(NRO&), 6), 33, 8)
                   Call REPLA(YR$, MKI$(FF%), 41, 2)
                   Call REPLA(YR$, MKD$(NETO# * SIG%), 43, 8)
                   CARESUVE& = CARESUVE& + 1
                   Call GRAREG("RESUVEN", YR$, CARESUVE&)
               End If
           End If
        End If
        '
    Next IL&
    '
    Call SETULTREG&("RESUVEN", CARESUVE&)
    Call CIERRARCH("RESUVEN")
    '
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    Call HEADERS("RESUVEN", "")
    Call HEADERS("RESUVEN", "Periodo: " + PERIO$)
    '
    Screen.MousePointer = 1
    Call FINAL("*RESUSOR")
    '
End Sub
'


