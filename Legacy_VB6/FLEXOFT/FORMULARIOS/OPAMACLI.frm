VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPAMACLI 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Clientes - Configuración"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   64
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   875
      Left            =   8925
      ScaleHeight     =   810
      ScaleWidth      =   60
      TabIndex        =   62
      Top             =   340
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "I.B.Pcia.Bs.As"
      Height          =   950
      Left            =   5040
      TabIndex        =   59
      Top             =   270
      Visible         =   0   'False
      Width           =   1800
      Begin VB.CommandButton Command4 
         Height          =   540
         Left            =   945
         Picture         =   "OPAMACLI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Listado Clientes para I.Brutos"
         Top             =   280
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Height          =   540
         Left            =   210
         Picture         =   "OPAMACLI.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
         Top             =   280
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Valores Preteterminados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   5040
      TabIndex        =   27
      Top             =   1470
      Width           =   6000
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1365
         TabIndex        =   55
         Text            =   " "
         Top             =   630
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   2205
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   630
         Width           =   3615
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1890
         TabIndex        =   53
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   51
         Text            =   " "
         Top             =   3255
         Width           =   2250
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   6
         Left            =   1365
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   3630
         Width           =   4455
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   5600
         TabIndex        =   49
         Top             =   3255
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1890
         TabIndex        =   42
         Top             =   2730
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1890
         TabIndex        =   41
         Top             =   2310
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1890
         TabIndex        =   40
         Top             =   1890
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1890
         TabIndex        =   39
         Top             =   1050
         Width           =   175
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   2205
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   2730
         Width           =   3615
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1365
         TabIndex        =   37
         Text            =   " "
         Top             =   2730
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1365
         TabIndex        =   36
         Text            =   " "
         Top             =   1890
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2205
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1050
         Width           =   3615
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1365
         TabIndex        =   34
         Text            =   " "
         Top             =   1050
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   1365
         TabIndex        =   33
         Text            =   " "
         Top             =   2310
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2205
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1890
         Width           =   3615
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2205
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   2310
         Width           =   3615
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2205
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   1470
         Width           =   3615
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   29
         Text            =   "  "
         Top             =   1470
         Width           =   570
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1890
         TabIndex        =   28
         Top             =   1470
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Madre:"
         Height          =   330
         Left            =   1365
         TabIndex        =   52
         Top             =   3360
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precios:"
         Height          =   225
         Left            =   -420
         TabIndex        =   48
         Top             =   1995
         Width           =   1710
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Cliente:"
         Height          =   225
         Left            =   -105
         TabIndex        =   47
         Top             =   735
         Width           =   1380
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría I-B"
         Height          =   330
         Left            =   -105
         TabIndex        =   46
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cond.Pago:"
         Height          =   225
         Left            =   -420
         TabIndex        =   45
         Top             =   2415
         Width           =   1710
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   330
         Left            =   -630
         TabIndex        =   44
         Top             =   2835
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         Height          =   225
         Left            =   105
         TabIndex        =   43
         Top             =   1575
         Width           =   1185
      End
   End
   Begin VB.TextBox COTEXT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   6300
      TabIndex        =   58
      Text            =   " "
      Top             =   2100
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   7140
      TabIndex        =   57
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2100
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   6825
      TabIndex        =   56
      Top             =   2100
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   855
      Left            =   9240
      Picture         =   "OPAMACLI.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   340
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Controles Operativos"
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
      TabIndex        =   21
      Top             =   3885
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Alta de Nueva Cuenta"
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
         Left            =   210
         TabIndex        =   22
         Top             =   420
         Width           =   4215
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3360
            TabIndex        =   25
            Top             =   735
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Primer Número Libre Mayor que ..."
            Height          =   750
            Left            =   1890
            TabIndex        =   24
            Top             =   315
            Width           =   1485
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Número Siguiente al Ultimo"
            Height          =   750
            Left            =   315
            TabIndex        =   23
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      Height          =   855
      Left            =   10185
      Picture         =   "OPAMACLI.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   340
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.CommandButton Command20 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3255
         TabIndex        =   1
         Top             =   525
         Width           =   175
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Opciones de Trabajo"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1680
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Clientes"
            Height          =   450
            Index           =   9
            Left            =   2310
            TabIndex        =   17
            Top             =   1050
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Cuentas Contables"
            Height          =   450
            Index           =   8
            Left            =   2310
            TabIndex        =   16
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Listas de Precios"
            Height          =   345
            Index           =   7
            Left            =   315
            TabIndex        =   12
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Condiciones de Pago"
            Height          =   450
            Index           =   6
            Left            =   315
            TabIndex        =   13
            Top             =   630
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Provincias y Zonas de Venta"
            Height          =   450
            Index           =   5
            Left            =   315
            TabIndex        =   14
            Top             =   1050
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Vendedores"
            Height          =   345
            Index           =   4
            Left            =   2310
            TabIndex        =   15
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         Height          =   880
         Left            =   3570
         Picture         =   "OPAMACLI.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3225
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         TabIndex        =   9
         Text            =   " "
         Top             =   525
         Width           =   1695
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
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command21 
            Caption         =   "Análisis A-B-C"
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Estadísticas"
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Críticos"
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Consignación"
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Negativos"
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Inventarios"
            Height          =   435
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -735
         TabIndex        =   20
         Top             =   630
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPAMACLI.frx":120A
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "OPAMACLI.frx":1284
      Top             =   1200
   End
End
Attribute VB_Name = "OPAMACLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command16_Click()
    '
    Command16.Enabled = False
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                x$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(x$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", x$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command2_Click()
   OPXX$ = "-1"
   If Option3.Value = True Then OPXX$ = TRIM$(Str$(Val(Text3.TEXT)))
   Call GRACONTROL("", "PRONUCLI", OPXX$)
   '
   For Index = 0 To 7
       TV11% = TIPFIELD(ARCHE$(Index), 1)
       LF11% = LENFIELD(ARCHE$(Index), 1)
       ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
       VAPRE$ = "": ECO$ = ""
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
             VAPRE$ = TRIM$(COTEXT(Index).TEXT)
          End If
       End If
       COTEXT(Index).TEXT = VAPRE$
       DETEXT(Index).TEXT = ECO$
       Call GRACONTROL("AMACLI", ARCHE$(Index), VAPRE$)
   Next Index
   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
     VDEF$ = String$(16, 0) + Space$(16)
     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-A"))), 1, 4)
     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-B"))), 5, 4)
     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-EXP"))), 9, 4)
     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "MINIPER"))), 13, 4)
     Call REPLA(VDEF$, CONTROL$("PERIBRU", "CUECONTA"), 17, 12)
     '
     OBX$ = OBJPLA$("PARPERIBRU", VDEF$)
     If OBX$ <> "" Then
       Call GRACONTROL$("PERIBRU", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
       Call GRACONTROL$("PERIBRU", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
       Call GRACONTROL$("PERIBRU", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
       Call GRACONTROL$("PERIBRU", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
       Call GRACONTROL$("PERIBRU", "CUECONTA", Mid$(OBX$, 17, 12))
     End If
     '
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call FILESORT("DAPIBCLI", "DAPIBCLJ", 1, 1, "ASC")
   Screen.MousePointer = 11
   FIN& = ULTREG&("DAPIBCLJ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DAPIBCLJ", U&)
     NCLIX% = CVI(Left$(XX$, 2))
     CATIBX% = Val(CATIBRU$(NCLIX%))
     NIBR$ = NIBRCLI$(NCLIX%)
     Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
     Call REPLA(XX$, NIBR$, 67, 16)
     If CATIBX% > 0 Then
       If TRIM$(NIBR$) = "" Then
         Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
       End If
     End If
     Call GRAREG("DAPIBCLJ", XX$, U&)
   Next U&
   Call CIERRARCH("DAPIBCLJ")
   Screen.MousePointer = 1
   '
   Call FINAL("*LIDAPIBC")
   '
   Command4.Enabled = True
End Sub

'
Private Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CARTAVALI
  Screen.MousePointer = 1
  '
  Call APLICACIONSKINS(Me)
  '
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "*UNIMED"
  TABLA$(2) = "TADEPOSI"
  TABLA$(3) = ""       ' grupos contables por articulo
  TABLA$(4) = "*VENDEDOR"
  TABLA$(5) = "TAPROVIN"
  TABLA$(6) = "CONPAGO"
  TABLA$(7) = "LISPRE02"
      If EXISTE%("LISPRE02.EXE") < 1 Then
        TABLA$(7) = "ADMIPRE"
      End If
  TABLA$(8) = "CUEINGAS"
  TABLA$(9) = "GRUCOCLI"
  '
  PRONUCLI$ = CONTROL$("", "PRONUCLI")
  If Val(PRONUCLI$) >= 0 Then
    Option3.Value = True
    Text3 = PRONUCLI$
  End If
  '
  For Index = 0 To 7
    VAPRE$ = TRIM$(CONTROL$("AMACLI", ARCHE$(Index)))
    COTEXT(Index) = "": DETEXT(Index) = ""
    If VAPRE$ <> "" Then
      TV11% = TIPFIELD(ARCHE$(Index), 1)
      LF11% = LENFIELD(ARCHE$(Index), 1)
      ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
      If ECO$ <> "" Then
          COTEXT(Index).TEXT = VAPRE$
          DETEXT(Index).TEXT = ECO$
        Else
          COTEXT(Index).TEXT = ""
          DETEXT(Index).TEXT = ""
       End If
    End If
    DETEXT(Index).Refresh
    DETEXT(Index).TabStop = False
  Next Index
  '
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then
    Frame3.Visible = True
  End If
  '
End Sub
'
Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For K% = 4 To 9
       If K% <> Index Then
          Option1(K%).Value = False
       End If
    Next K%
    Command16.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOCLI")
            Call CIERRARCH("!GRUCOCLI")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOCLI", "")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                x$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(x$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", x$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
End Sub


Private Sub Option2_Click()
   Option2.Value = True
   Text3.TEXT = ""
End Sub

Private Sub Option3_Click()
   Option3.Value = True
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text3_Click()
   Option3.Value = True
End Sub



Private Sub COTEXT_DblClick(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO(ARCHE$(Index))
    If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
    COTEXT(Index).SetFocus
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub


