VERSION 5.00
Begin VB.Form SETPROMA07 
   BackColor       =   &H00F0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Proveedores - Configuración"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5940
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
   ScaleHeight     =   5100
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404000&
      ForeColor       =   &H00404000&
      Height          =   5600
      Left            =   5055
      ScaleHeight     =   5535
      ScaleWidth      =   1440
      TabIndex        =   56
      Top             =   -360
      Width           =   1500
      Begin VB.Frame Frame3 
         BackColor       =   &H00404000&
         Caption         =   "Gan"
         ForeColor       =   &H8000000E&
         Height          =   810
         Left            =   150
         TabIndex        =   68
         Top             =   7000
         Visible         =   0   'False
         Width           =   600
         Begin VB.CommandButton Command5 
            Height          =   600
            Left            =   10
            Picture         =   "SETPROMA07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Configuracion Funcionamiento Retenciones de Ganancias"
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   60
         Top             =   4560
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   5000
            Left            =   -420
            TabIndex        =   61
            Top             =   -105
            Width           =   12000
         End
      End
      Begin VB.Frame c 
         BackColor       =   &H00404000&
         Caption         =   "I-Br"
         ForeColor       =   &H8000000E&
         Height          =   1445
         Left            =   150
         TabIndex        =   59
         Top             =   7000
         Width           =   600
         Begin VB.CommandButton Command3 
            Height          =   600
            Left            =   10
            Picture         =   "SETPROMA07.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Configuracion Funcionamiento Percepciones IB Pcia de Bs.As"
            Top             =   210
            Width           =   600
         End
         Begin VB.CommandButton Command4 
            Height          =   600
            Left            =   10
            Picture         =   "SETPROMA07.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Listado Proveedores para I.Brutos"
            Top             =   840
            Width           =   600
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   600
         Left            =   150
         Picture         =   "SETPROMA07.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   1170
         Width           =   600
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         Height          =   600
         Left            =   150
         Picture         =   "SETPROMA07.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   540
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "SETPROMA07.frx":1342
         Top             =   4800
         Width           =   1515
      End
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
      Left            =   7470
      TabIndex        =   51
      Top             =   5460
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
      Index           =   7
      Left            =   6945
      TabIndex        =   52
      Text            =   " "
      Top             =   5460
      Width           =   570
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F0FFD0&
      Caption         =   "Valores Preteterminados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4950
      Left            =   240
      TabIndex        =   23
      Top             =   5160
      Width           =   4320
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
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   735
         Width           =   4035
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
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1575
         Width           =   4035
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
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2415
         Width           =   4035
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
         Height          =   285
         Index           =   6
         Left            =   3990
         TabIndex        =   31
         Top             =   3900
         Width           =   175
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
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3255
         Width           =   4035
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
         Left            =   3990
         TabIndex        =   53
         Top             =   2835
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
         Index           =   4
         Left            =   3465
         TabIndex        =   54
         Text            =   " "
         Top             =   2835
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
         Left            =   3990
         TabIndex        =   24
         Top             =   1995
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
         Left            =   1785
         TabIndex        =   33
         Text            =   " "
         Top             =   3900
         Width           =   2250
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
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4320
         Width           =   4035
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
         Left            =   3990
         TabIndex        =   27
         Top             =   1155
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
         Index           =   2
         Left            =   3465
         TabIndex        =   26
         Text            =   " "
         Top             =   1155
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
         Index           =   0
         Left            =   3465
         TabIndex        =   25
         Text            =   "  "
         Top             =   1995
         Width           =   570
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Pago:"
         Height          =   225
         Left            =   525
         TabIndex        =   55
         Top             =   2940
         Width           =   2865
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5985
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Madre:"
         Height          =   330
         Left            =   -210
         TabIndex        =   34
         Top             =   4005
         Width           =   1905
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Proveedor:"
         Height          =   225
         Left            =   1050
         TabIndex        =   30
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría de Ingresos Brutos:"
         Height          =   330
         Left            =   630
         TabIndex        =   29
         Top             =   1260
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         Height          =   225
         Left            =   2205
         TabIndex        =   28
         Top             =   2100
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
      Left            =   6090
      TabIndex        =   37
      Text            =   " "
      Top             =   5385
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
      Left            =   5775
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   7065
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
      Left            =   6615
      TabIndex        =   35
      Top             =   5385
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0FFD0&
      Caption         =   "Controles Operativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   105
      TabIndex        =   18
      Top             =   2520
      Width           =   4750
      Begin VB.Frame Frame12 
         BackColor       =   &H00F0FFD0&
         Caption         =   "Formato de Número de CUIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   160
         TabIndex        =   70
         Top             =   1560
         Width           =   4430
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Suprime Validación de Formato (Argentina)"
            Height          =   255
            Left            =   105
            TabIndex        =   71
            Top             =   300
            Width           =   3960
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F0FFD0&
         Caption         =   "Alta de Nueva Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   160
         TabIndex        =   19
         Top             =   300
         Width           =   4425
         Begin VB.Frame Frame6 
            BackColor       =   &H00F0FFD0&
            Caption         =   "PRF"
            Height          =   750
            Left            =   105
            TabIndex        =   41
            Top             =   315
            Width           =   855
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   105
               TabIndex        =   42
               Top             =   315
               Width           =   620
            End
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3570
            TabIndex        =   22
            Top             =   735
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Primer Número Libre Ma-         yor que ..."
            Height          =   750
            Left            =   2415
            TabIndex        =   21
            Top             =   315
            Width           =   1905
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Número Siguiente al Ultimo"
            Height          =   750
            Left            =   1155
            TabIndex        =   20
            Top             =   315
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00F0FFD0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2210
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4750
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
         Height          =   285
         Left            =   3570
         TabIndex        =   1
         Top             =   285
         Width           =   175
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00F0FFD0&
         Caption         =   "Opciones de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   160
         TabIndex        =   16
         Top             =   1125
         Width           =   4425
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Grupos de Proveedores"
            Height          =   450
            Index           =   9
            Left            =   2520
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Cuentas Contables"
            Height          =   450
            Index           =   8
            Left            =   2520
            TabIndex        =   14
            Top             =   1500
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Condiciones de Pago"
            Height          =   450
            Index           =   6
            Left            =   420
            TabIndex        =   12
            Top             =   315
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0FFD0&
            Caption         =   "Provincias y Zonas de Venta"
            Height          =   450
            Index           =   5
            Left            =   2400
            TabIndex        =   13
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
         Height          =   700
         Left            =   3885
         Picture         =   "SETPROMA07.frx":3264
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   690
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
         Height          =   285
         Left            =   160
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3590
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
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   285
         Width           =   1905
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
         Height          =   225
         Left            =   -735
         TabIndex        =   17
         Top             =   390
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "NO VISIBLES"
      Height          =   1170
      Left            =   4440
      TabIndex        =   38
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
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
         Left            =   2835
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   0
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
         Index           =   3
         Left            =   1995
         TabIndex        =   47
         Text            =   " "
         Top             =   0
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
         Index           =   5
         Left            =   1995
         TabIndex        =   46
         Text            =   " "
         Top             =   840
         Width           =   570
      End
      Begin VB.TextBox DETEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   2835
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "  "
         Top             =   840
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
         Index           =   3
         Left            =   2520
         TabIndex        =   44
         Top             =   0
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
         Left            =   2520
         TabIndex        =   43
         Top             =   840
         Width           =   175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Vendedores"
         Height          =   345
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   420
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listas de Precios"
         Height          =   345
         Index           =   7
         Left            =   0
         TabIndex        =   39
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   330
         Left            =   0
         TabIndex        =   50
         Top             =   945
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precios:"
         Height          =   225
         Left            =   210
         TabIndex        =   49
         Top             =   105
         Width           =   1710
      End
   End
End
Attribute VB_Name = "SETPROMA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

 Sub Command16_Click()
    '
    Command16.Enabled = False
    Call COPYSTRU("GRUCOCLI", "GRUCOPRO")
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOPRO" Then
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOPRO")
            Call CIERRARCH("!GRUCOPRO")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOPRO", "")
            VTB% = VENTABU%("GRUCOPRO")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOPRO")
                X$ = REGLEIDO$("!GRUCOPRO", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOPRO", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOPRO", J&)
            Call CIERRARCH("GRUCOPRO")
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
'ANTIGUO CODIGO
'Command16.Enabled = False
'    If TRIM$(MODUCONE$) <> "" Then
'        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
'            LDT$ = LUDAT$
'            Call CIERRARCH("GRUCOCLI")
'            Call CIERRARCH("!GRUCOCLI")
'            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
'            Call TRALOCAL("GRUCOCLI", "")
'            VTB% = VENTABU%("GRUCOCLI")
'            If VTB% < 0 Then GoTo 99
'            J& = 0
'            For I& = 1 To ULTREG&("!GRUCOCLI")
'                X$ = REGLEIDO$("!GRUCOCLI", I&)
'                If CVI(Left$(X$, 2)) > 0 Then
'                    J& = J& + 1
'                    Call GRAREG("GRUCOCLI", X$, J&)
'                End If
'            Next I&
'            Call SETULTREG("GRUCOCLI", J&)
'            Call CIERRARCH("GRUCOCLI")
'          ElseIf MODUCONE = "TADEPOSI" Then
'            Call CARTADEPOSI
'          ElseIf MODUCONE = "TATIMAT" Then
'            Call CARTATIMA
'          ElseIf MODUCONE = "CONPAGO" Then
'            Call CARCONPAGOS
'          ElseIf MODUCONE = "TAPROVIN" Then
'            Call CARTAPROVIN
'          ElseIf MODUCONE = "CUEINGAS" Then
'            Call CARCUEINGAS
'          Else
'            Call CONECRUN(MODUCONE$, Text2.TEXT)
'        End If
'    End If
'99  Command16.Enabled = True
End Sub

Private Sub Command2_Click()
   OPXX$ = "-1"
   If Option3.Value = True Then OPXX$ = TRIM$(Str$(Val(Text3.TEXT)))
   Call GRACONTROL("", "PRONUPRO", OPXX$)
   '
   '\\\OT-05-0198-WAR-19/04/05-///
   VALICUIT1$ = "NO"
   If Check1.Value = 1 Then VALICUIT1$ = "SI"
   Call GRACONTROL("FORMCUIT", "SUPRIVAL", VALICUIT1$)
   '\\\OT-05-0198-WAR-FIN///
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
       Call GRACONTROL("AMAPRO", ARCHE$(Index), VAPRE$)
   Next Index
   '
End Sub



Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACPAIBRU", "Actualizar Parámetros Ingresos Brutos") = 2 Then
     If ULTREG&("TACATIBP") < 5 Then
       Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
       Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
       Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
       Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
       Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
       Call CIERRARCH("TACATIBP")
     End If
     '
     Call TRALOCAL("TACATIBP", "")
     VTB% = VENTABU%("TACATIBP")
     If VTB% >= 0 Then
       For u& = 1 To ULTREG&("!TACATIBP")
         XX$ = REGLEIDO$("!TACATIBP", u&)
         Call REPLA(XX$, Chr$(u& - 1), 1, 1)
         Call GRAREG("TACATIBP", XX$, u&)
       Next u&
       Call CIERRARCH("TACATIBP")
     End If
   End If
   Command3.Enabled = True

'   Command3.Enabled = False
'   If DERACCE%("ACPAIBRU", "Actualizar Parametros Ingresos Brutos") = 2 Then
'     VDEF$ = String$(16, 0) + Space$(16)
'     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-A"))), 1, 4)
'     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-B"))), 5, 4)
'     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "TASA-EXP"))), 9, 4)
'     Call REPLA(VDEF$, MKS$(Val(CONTROL$("PERIBRU", "MINIPER"))), 13, 4)
'     Call REPLA(VDEF$, CONTROL$("PERIBRU", "CUECONTA"), 17, 12)
'     '
'     OBX$ = OBJPLA$("PARPERIBRU", VDEF$)
'     If OBX$ <> "" Then
'       Call GRACONTROL$("PERIBRU", "TASA-A", Str(CVS(Mid$(OBX$, 1, 4))))
'       Call GRACONTROL$("PERIBRU", "TASA-B", Str(CVS(Mid$(OBX$, 5, 4))))
'       Call GRACONTROL$("PERIBRU", "TASA-EXP", Str(CVS(Mid$(OBX$, 9, 4))))
'       Call GRACONTROL$("PERIBRU", "MINIPER", Str(CVS(Mid$(OBX$, 13, 4))))
'       Call GRACONTROL$("PERIBRU", "CUECONTA", Mid$(OBX$, 17, 12))
'     End If
'     '
'   End If
'   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call FILESORT("DARIBPRO", "DARIBPRJ", 1, 1, "ASC")
   Screen.MousePointer = 11
   FIN& = ULTREG&("DARIBPRJ")
   For u& = 1 To FIN&
     Call TRACE(20, u&, FIN&)
     XX$ = REGLEIDO$("DARIBPRJ", u&)
     NCLIX% = CVI(Left$(XX$, 2))
     CATIBX% = Val(CATIBRU$((-1) * NCLIX%))
     NIBR$ = NIBRCLI$((-1) * NCLIX%)
     Call REPLA(XX$, MKI$(CATIBX%), 65, 2)
     Call REPLA(XX$, NIBR$, 67, 16)
     If CATIBX% > 0 Then
       If TRIM$(NIBR$) = "" Then
         Call REPLA(XX$, "Falta Nro. Inscripción", 83, 24)
       End If
     End If
     Call GRAREG("DARIBPRJ", XX$, u&)
   Next u&
   Call CIERRARCH("DARIBPRJ")
   Screen.MousePointer = 1
   '
   Call FINAL("*LIDARIBP")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  Command5.Enabled = False
  '\\\OT-05-0635-WAR-22/09/05///
  RTA% = ALTERNA%("Tabla de Categorias\Escalas de Retención")
  If RTA% = 1 Then
      Call CARTARET
    ElseIf RTA% = 2 Then
      Call ESCALRET
  End If
  Command5.Enabled = True
  '\\\OT-05-0635-WAR-22/09/05///
End Sub
Sub ESCALRET() 'ESCALAS DE RETENCION DE GANANCIAS
  '
  Call SELECHO("TACRETG")
  If TRIM$(VALACT1$("TACRETG")) <> "" Then
    '3.- SI NO TIENE MINIMO O TASA => AFUERA
    NUCA% = Val(VALACT1$("TACRETG")) 'CATEGORIA ELEGIDA
    X$ = FILTERGETRECORD$("TACRETG", 1, Chr(NUCA%))
    DESCRICAT$ = Mid$(X$, 2, 31)
    MINEX1 = CVD(Mid$(X$, 33, 8))
    TASA1 = CVS(Mid$(X$, 41, 4))
    If MINEX1 <= 0.005 And TASA1 <= 0.005 Then
      Call COMUNI("No se puede Cargar Escala a Categoria Exento")
      Exit Sub
    End If
    '6.- CARGA LA TABLA C/REGISTROS Q' MACHEEN CON LA CATEGORIA SELECIONADA
    Call FILTERFILE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    '7.-GRABACION
    VTB% = VENTABU%("ESCALGAN/NC/TITUPAN=Escala de Retencion - " & DESCRICAT$ & "")
    If VTB% >= 0 Then
      For J& = 1 To ULTREG&("!ESCALGAN")
        YY$ = REGLEIDO$("!ESCALGAN", J&)
        Call REPLA(YY$, MKI$(NUCA%), 1, 2)
        Call GRAREG("!ESCALGAN", YY$, J&)
      Next J&
      Call FILTERUPDATE("ESCALGAN", "!ESCALGAN", 1, MKI$(NUCA%))
    End If
    '
  End If
  '
End Sub
Sub CARTARET()
    Call TRALOCAL("TACRETG", "")
     VTB% = VENTABU%("TACRETG")
     If VTB% >= 0 Then
       For u& = 1 To ULTREG&("!TACRETG")
         XX$ = REGLEIDO$("!TACRETG", u&)
         ' Graba el numero de orden automaticamente
         Call REPLA(XX$, Chr$(u& - 1), 1, 1)
         Call GRAREG("TACRETG", XX$, u&)
       Next u&
       Call CIERRARCH("TACRETG")
     End If
End Sub

'
Private Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CARTAVALI
  Screen.MousePointer = 1
  '
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "*UNIMED"
  TABLA$(2) = "TADEPOSI"
  TABLA$(3) = ""       ' grupos contables por articulo
  TABLA$(4) = "*VENDEDOR"
  TABLA$(5) = "TAPROVIN"
  TABLA$(6) = "CONPAGO"
  TABLA$(7) = "LISPRE04"
  TABLA$(8) = "CUEINGAS"
  '
  TABLA(9) = "GRUCOPRO" 'NUEVO
  'TABLA$(9) = "GRUCOCLI"
  
  '
  PRONUCLI$ = CONTROL$("", "PRONUPRO")
  If Val(PRONUCLI$) >= 0 Then
    Option3.Value = True
    Text3 = PRONUCLI$
  End If
  '
  For Index = 0 To 7
    VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
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
  Frame3.Visible = True
  ' '\\\OT-05-0198-WAR-19/04/05-///
  If CONTROL$("FORMCUIT", "SUPRIVAL") = "SI" Then
      Check1.Value = 1
  End If
  '''\\\OT-05-0198-WAR-FIN///
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
    On Error Resume Next
    Command16.SetFocus
    On Error GoTo 0
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOPRO" Then
            Call COPYSTRU("GRUCOCLI", "GRUCOPRO") 'NUEVO
            LDT$ = LUDAT$
            Call CIERRARCH("GRUCOPRO")
            Call CIERRARCH("!GRUCOPRO")
            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            Call TRALOCAL("GRUCOPRO", "")
            VTB% = VENTABU%("GRUCOPRO")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOPRO")
                X$ = REGLEIDO$("!GRUCOPRO", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOPRO", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOPRO", J&)
            Call CIERRARCH("GRUCOPRO")
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
    'CODIGO ANTIGUO
'       If TRIM$(MODUCONE$) <> "" Then
'        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
'            LDT$ = LUDAT$
'            Call CIERRARCH("GRUCOCLI")
'            Call CIERRARCH("!GRUCOCLI")
'            'Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
'            Call TRALOCAL("GRUCOCLI", "")
'            VTB% = VENTABU%("GRUCOCLI")
'            If VTB% < 0 Then Exit Sub
'            J& = 0
'            For I& = 1 To ULTREG&("!GRUCOCLI")
'                X$ = REGLEIDO$("!GRUCOCLI", I&)
'                If CVI(Left$(X$, 2)) > 0 Then
'                    J& = J& + 1
'                    Call GRAREG("GRUCOCLI", X$, J&)
'                End If
'            Next I&
'            Call SETULTREG("GRUCOCLI", J&)
'            Call CIERRARCH("GRUCOCLI")
'          ElseIf MODUCONE = "TADEPOSI" Then
'            Call CARTADEPOSI
'          ElseIf MODUCONE = "TATIMAT" Then
'            Call CARTATIMA
'          ElseIf MODUCONE = "CONPAGO" Then
'            Call CARCONPAGOS
'          ElseIf MODUCONE = "TAPROVIN" Then
'            Call CARTAPROVIN
'          ElseIf MODUCONE = "CUEINGAS" Then
'            Call CARCUEINGAS
'          Else
'            Call CONECRUN(MODUCONE$, Text2.TEXT)
'        End If
'    End If
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


