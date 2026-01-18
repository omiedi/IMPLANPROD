VERSION 5.00
Begin VB.Form COGECOMP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Gestión de Compras - Cuentas por Pagar"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command16 
      Caption         =   "Inf. CITI"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9765
      Picture         =   "COGECOMP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Resúmenes Estadísticos de Compras"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Utilities"
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
      Left            =   9765
      Picture         =   "COGECOMP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Utilitarios - Descarga de Movimientos"
      Top             =   5250
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9750
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   55
      Top             =   2205
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4515
      Sorted          =   -1  'True
      TabIndex        =   54
      Top             =   3465
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0E0&
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
      Left            =   5145
      TabIndex        =   44
      Top             =   210
      Width           =   4425
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   45
         Top             =   1260
         Width           =   2325
         Begin VB.CommandButton Command13 
            Enabled         =   0   'False
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
            Picture         =   "COGECOMP.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Detalle de Compras por Proveedor - Rubro - Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalles por  Pro- veedor y Rubro"
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
            TabIndex        =   47
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
         TabIndex        =   48
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command14 
            Enabled         =   0   'False
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
            Picture         =   "COGECOMP.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Detalle de Compras en Pesos por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle por Pro- veedor en Pesos"
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
            TabIndex        =   50
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
         TabIndex        =   51
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command15 
            Enabled         =   0   'False
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
            Picture         =   "COGECOMP.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Estadisticas de Compras"
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
            TabIndex        =   53
            Top             =   0
            Width           =   1410
         End
      End
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
      Height          =   855
      Left            =   9765
      Picture         =   "COGECOMP.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   4305
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0E0&
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
      Left            =   5145
      TabIndex        =   6
      Top             =   2415
      Width           =   4425
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0E0&
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
               Picture         =   "COGECOMP.frx":12EC
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Tablas de Validación de Datos"
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
               Picture         =   "COGECOMP.frx":15F6
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
         BackColor       =   &H00C0E0E0&
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
               Picture         =   "COGECOMP.frx":1900
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Acceso a Padrón de Proveedores y Subcontratistas"
               Top             =   0
               Width           =   560
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
               Left            =   735
               TabIndex        =   37
               Top             =   105
               Width           =   1485
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
               Picture         =   "COGECOMP.frx":1C0A
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Acceso al Maestro de Artículos de Stocks"
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
               Left            =   735
               TabIndex        =   34
               Top             =   0
               Width           =   1485
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
               Left            =   0
               Picture         =   "COGECOMP.frx":1F14
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Acceso a la Base de Cuentas Contables"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Base de Cuen- tas Contables"
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
               TabIndex        =   31
               Top             =   0
               Width           =   1485
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Compras y Acreedores"
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
      Left            =   315
      TabIndex        =   3
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0E0&
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
         Height          =   1380
         Left            =   420
         TabIndex        =   40
         Top             =   4935
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   41
            Top             =   500
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
               Picture         =   "COGECOMP.frx":221E
               Style           =   1  'Graphical
               TabIndex        =   42
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
               TabIndex        =   43
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0E0&
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
         Height          =   2955
         Left            =   420
         TabIndex        =   5
         Top             =   1840
         Width           =   3585
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0E0&
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
            Top             =   2100
            Width           =   3165
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0E0E0&
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
               BackColor       =   &H00C0E0E0&
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
               Picture         =   "COGECOMP.frx":2528
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Listado de Saldos y Deudas de Acreedores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldos y Deudas de Acreedores"
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
               Top             =   105
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
               Picture         =   "COGECOMP.frx":2832
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Listado de Planificación Financiera de Pagos"
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
               Picture         =   "COGECOMP.frx":2B3C
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Listados de Obligaciones Pendientes de Pago"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuentas por Pagar"
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
         BackColor       =   &H00C0E0E0&
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
         Height          =   1275
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   9
            Top             =   400
            Width           =   2325
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
               Picture         =   "COGECOMP.frx":2E46
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Proveedor"
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
      Height          =   855
      Left            =   9750
      Picture         =   "COGECOMP.frx":3150
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
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
      Left            =   9750
      Picture         =   "COGECOMP.frx":329A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1260
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
      Left            =   9765
      Picture         =   "COGECOMP.frx":35A4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9200
      Picture         =   "COGECOMP.frx":38AE
      Top             =   6210
      Width           =   2010
   End
End
Attribute VB_Name = "COGECOMP"
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
    Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    Command11.Enabled = True
End Sub


Private Sub Command12_Click()
    Command12.Enabled = False
    Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    Command12.Enabled = True
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CARCUEINGAS
    Command10.Enabled = True
End Sub

Private Sub Command13_Click()
    '
Call OPNOIN(""): Exit Sub
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
         'Call CALREVEPRO(DES%, HAS%)
    End If
    '
End Sub
'

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    Call CONECRUN("RESTACOM", "Resúmenes Estadísticos de Compras")
    Command15.Enabled = True
    '
End Sub

Private Sub Command14_Click()
    '
Call OPNOIN(""): Exit Sub
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
         'Call CALRESUCO(DES%, HAS%)
    End If
    '
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    '
    Dim IVAPROV#(32767)
    '
    LIMIVA# = 500
    '
    HOO% = HOY(HO$)
10  VDEF$ = MKI$(COMESACT%) + MKI$(HOO%)
    OPX$ = OBJPLA$("DESHASFECHA", VDEF$)
    '
    If OPX$ <> "" Then
      DES% = CVI(Left$(OPX$, 2))
      HAS% = CVI(Mid$(OPX$, 3, 2))
      If DES% > HAS% Then GoTo 10
      '
      J& = 0
      FIN& = ULTREG&("CUECORR")
      Screen.MousePointer = 11
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("CUECORR", U&)
        FECHMOV% = CVI(Mid$(X$, 1, 2))
        If FECHMOV% >= DES% Then
          If FECHMOV% <= HAS% Then
            If Mid$(X$, 7, 2) <> "NC" Then ' No incluye Notas de Credito
              NPX% = CVI(Mid$(X$, 5, 2))
              '\\\OT-5231-FG-17/05/05///
              'IMPDOC# = CVD(Mid$(X$, 75, 8)) + CVD(Mid$(X$, 83, 8)) + CVD(Mid$(X$, 91, 8))
              IMPDOC# = CVD(Mid$(X$, 75, 8)) + CVD(Mid$(X$, 91, 8)) ' No considera IVA Aduana
              '\\\OT-5231-FG-FIN///
              If IMPDOC# >= LIMIVA# Then
                COMPROB$ = Mid$(X$, 7, 18)
                CUIT$ = CUITCLI$((-1) * NPX%)
                '
                Y$ = REGBLAN$("TRANCITI")
                Call REPLA(Y$, COMPROB$, 1, 18)
                Call REPLA(Y$, MKI$(FECHMOV%), 19, 2)
                Call REPLA(Y$, CUIT$, 21, 13)
                Call REPLA(Y$, MKI$(NPX%), 34, 2)
                Call REPLA(Y$, MKD$(IMPDOC#), 36, 8)
                '
                J& = J& + 1
                Call GRAREG("TRANCITI", Y$, J&)
              Else
                IVAPROV#(NPX%) = IVAPROV#(NPX%) + IMPDOC#
              End If
            End If
          End If
        End If
      Next U&
      '
      Call SETULTREG("TRANCITI", J&)
      Call CIERRARCH("*.*")
      '
      ASALID$ = "C:\FLEXOFT\TRANCITI.TXT"
      NX% = FILEOPEN%(ASALID$, 2, 128)
      CAREGI% = 0
      '
      FIN& = ULTREG&("TRANCITI")
      Call TRACE(20, U&, FIN&)
      For U& = 1 To FIN&
        X$ = REGLEIDO$("TRANCITI", U&)
        TIPO$ = Mid$(X$, 1, 2)
        Select Case TIPO$
            Case "FC"
                 TIPOCOM$ = "01"
            Case "ND"
                 TIPOCOM$ = "02"
            Case "NC"
                 TIPOCOM$ = "03"
            Case "TF"
                 TIPOCOM$ = "81"
            Case Else
                 TIPOCOM$ = "39"
        End Select
        YY$ = Space(130)
        NPX% = CVI(Mid$(X$, 34, 2))
        Call REPLA(YY$, TIPOCOM$, 1, 2)
          If TIPOCOM$ = 39 Then
             COMPROB$ = TRIM$(Mid$(X$, 6, 13))
             While Len(COMPROB$) < 24: COMPROB$ = "0" + COMPROB$: Wend
           Else
             COMPROB$ = Mid$(X$, 6, 4) & String$(12, "0") & Mid$(X$, 11, 8)
          End If
        Call REPLA(YY$, COMPROB$, 3, 24)
          FECHAA$ = FECHATEX$(CVI(Mid$(X$, 19, 2)))
          FECHA$ = Mid$(FECHAA$, 1, 2) & Mid$(FECHAA$, 4, 2) & "20" & Mid$(FECHAA$, 7, 4)
        Call REPLA(YY$, FECHA$, 27, 8)
          CUIT$ = REPLACAR$(Mid$(X$, 21, 13), "-", "")
        Call REPLA(YY$, CUIT$, 35, 11)
          RASO$ = Mid$(RASOCLI$((-1) * NPX%), 1, 25)
        Call REPLA(YY$, RASO$, 46, 25)
          IMPORTE$ = TRIM$(FORMATNUM$(100 * CVD(Mid$(X$, 36, 8)), "F12.0"))
          Do While Len(IMPORTE$) < 12
            IMPORTE$ = "0" & IMPORTE$
          Loop
        Call REPLA(YY$, IMPORTE$, 71, 12)
        Call REPLA(YY$, String$(11, "0"), 83, 93)
        Call REPLA(YY$, String$(12, "0"), 119, 130)
        '
        Print #NX%, YY$
        CAREGI% = CAREGI% + 1
      Next U&
      '
      FIN& = 32767
      For U& = 1 To 32767
        Call TRACE(20, U&, FIN&)
        If Abs(IVAPROV#(U&)) >= 0.005 Then
          YY$ = Space(130)
          NPX% = U&
          Call REPLA(YY$, String$(34, "0"), 1, 34)
            CUIT$ = REPLACAR$(CUITCLI$((-1) * NPX%), "-", "")
          Call REPLA(YY$, CUIT$, 35, 11)
            RASO$ = Mid$(RASOCLI$((-1) * NPX%), 1, 25)
          Call REPLA(YY$, RASO$, 46, 25)
            IMPORTE$ = TRIM$(FORMATNUM$(100 * IVAPROV#(U&), "F12.0"))
            Do While Len(IMPORTE$) < 12
              IMPORTE$ = "0" & IMPORTE$
            Loop
          Call REPLA(YY$, IMPORTE$, 71, 12)
          Call REPLA(YY$, String$(11, "0"), 83, 93)
          Call REPLA(YY$, String$(12, "0"), 119, 130)
          '
          Print #NX%, YY$
          CAREGI% = CAREGI% + 1
        End If
      Next U&
      '
      Close #NX%
      ASALID1$ = REPLACAR$(ASALID$, "\", "/")
      Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
      '
      Screen.MousePointer = 1
      '
    End If
99  Call CIERRARCH("*.*")
    Command16.Enabled = True
    '
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    Call COMUNI("Opción no Accesible Desde esta Pantalla.\  \Acceda Desde el Escritorio de A-B-M Proveedores.")
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   Call COMUNI("No hay Opciones de Configuración\para la Presente Aplicación.")
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   '
   ' PARA RECUPERAR MOVIMIENTOS AL 23 ABRIL DE 2005
   Stop
   
   Call COPYSTRU("CUECORR", "CUECORVI")
   Call COPYSTRU("CUECORR", "CUECORNU")
   Call COPYSTRU("CUECOSS", "CUECOSVI")
   Call COPYSTRU("CUECOSS", "CUECOSNU")
   '
   JJ& = 0
   For U& = 1 To 21699
     XX$ = REGLEIDO$("CUECORVI", U&)
     YY$ = REGLEIDO$("CUECOSVI", U&)
     JJ& = JJ& + 1
     Call GRAREG("CUECORR", XX$, JJ&)
     Call GRAREG("CUECOSS", YY$, JJ&)
   Next U&
   '
   For U& = 21700 To 21708
     XX$ = REGLEIDO$("CUECORVI", U&)
     YY$ = REGBLAN$("CUECOSS")
     Call REPLA(YY$, XX$, 1, 300)
     JJ& = JJ& + 1
     Call GRAREG("CUECORR", XX$, JJ&)
     Call GRAREG("CUECOSS", YY$, JJ&)
   Next U&
   '
   For U& = 21707 To ULTREG&("CUECORNU")
     XX$ = REGLEIDO$("CUECORNU", U&)
     YY$ = REGLEIDO$("CUECOSNU", U&)
     JJ& = JJ& + 1
     Call GRAREG("CUECORR", XX$, JJ&)
     Call GRAREG("CUECOSS", YY$, JJ&)
   Next U&
   '
   Call SETULTREG("CUECORR", JJ&)
   Call SETULTREG("CUECOSS", JJ&)
   Call CIERRARCH("CUECORR")
   Call CIERRARCH("CUECOSS")
   Stop
   '
''   AA1 = ULTREG&("CUECORVI")
''   AA2 = ULTREG&("CUECOSVI")
''   '
''   For U& = 21708 To 21650 Step -1
''     XX$ = REGLEIDO$("CUECORVI", U&)
''     YY$ = REGLEIDO$("CUECORNU", U& - 2)
''     MsgBox Mid$(XX$, 7, 18) & "   " & Mid$(YY$, 7, 18)
''   Next U&
''   UVI& = AA1
''   UREVI$ = REGLEIDO$("CUECORVI", UVI&)
''   For U2& = ULTREG&("CUECORNU") To 1 Step -1
''     YY$ = REGLEIDO$("CUECORNU", U2&)
''     If YY$ = UREVI$ Then
''       AA3 = AA1
''     End If
''   Next U2&
   '
End Sub

'
Private Sub Command2_Click()
    CARGAMEN% = 0
    REMECOMP.Show 1
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command3_Click()
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
    If ULTREG&("ACREDOR") > 0 Then
        X$ = REGLEIDO$("ACREDOR", 1)      ' fuerza la
        Call GRAREG("ACREDOR", X$, 1)     ' reindexacion
        Call CIERRARCH("ACREDOR")         ' del archivo
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("ACREDOR")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    FAPEPAGO.Text10.TEXT = Str$(NCMIN%)
    FAPEPAGO.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAGO.Show 1
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call OPNOIN("Planificacion Finaciera de Pagos")
    '
    Screen.MousePointer = 11
    Dim SALDEUPEN#(32768)
    '
    If ULTREG&("ACREDOR") > 0 Then
        X$ = REGLEIDO$("ACREDOR", 1)      ' fuerza la
        Call GRAREG("ACREDOR", X$, 1)     ' reindexacion
        Call CIERRARCH("ACREDOR")         ' del archivo
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("ACREDOR")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    FAPEPAGO.Text10.TEXT = Str$(NCMIN%)
    FAPEPAGO.Text8.TEXT = Str$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
    CARFAPEN% = 0
    FAPEPAGO.Show 1
    '
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call GELISALPRO
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call CONECRUN("PRIFORMS", "Diseño de Formularios de Impresion")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Hide
    RESCUPRO.Show
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGECOM")
End Sub

Private Sub Command9_Click()
    DESCACUECO.Show 1
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "COGECOM"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(EPAC$, "VDR") > 0 Then
      Call GRACONTROL("", "INTREMOT", "")   ' 31/03/04 - borra integracion remota
      On Error Resume Next
      Kill "\FLEXOFT\VIEJOS\*.*"            ' BORRA VERSIONES VIEJAS
      On Error GoTo 0
    End If
    '
    If EXISTE%("RESTACOM.EXE") > 0 Then Command15.Enabled = True
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
