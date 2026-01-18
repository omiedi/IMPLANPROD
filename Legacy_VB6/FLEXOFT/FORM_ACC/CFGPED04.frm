VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CFGPED04 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Pedidos de Clientes "
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture6 
      Height          =   6370
      Left            =   5460
      ScaleHeight     =   6315
      ScaleWidth      =   1215
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      Begin VB.CommandButton Command23 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         Picture         =   "CFGPED04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
         Top             =   5400
         Width           =   650
      End
      Begin VB.Image Image7 
         Height          =   6450
         Left            =   0
         Picture         =   "CFGPED04.frx":030A
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2625
      TabIndex        =   58
      Text            =   " "
      Top             =   1260
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3990
      TabIndex        =   53
      Top             =   1260
      Width           =   175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Abrir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4305
      Picture         =   "CFGPED04.frx":26A2
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1260
      Width           =   750
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Importación de Pedidos"
      Height          =   1275
      Left            =   210
      TabIndex        =   45
      Top             =   6430
      Width           =   6000
      Begin VB.CommandButton Command11 
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
         Left            =   5670
         TabIndex        =   48
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command10 
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
         Left            =   5670
         TabIndex        =   46
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4725
         TabIndex        =   50
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Corte - Inicio de Importación:"
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
         Left            =   735
         TabIndex        =   49
         Top             =   915
         Width           =   3900
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   47
         Top             =   420
         Width           =   5580
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   5460
      ScaleHeight     =   6315
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command13 
         Caption         =   "Tablas"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Acceso a Tablas de Telas y Colores"
         Top             =   4095
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cotiz."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Parametros de Cálculo - Tipo de Cambio Dolar"
         Top             =   2535
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "S-Cod"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED04.frx":2C40
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Tabla de Sub-Codigos"
         Top             =   3150
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Left            =   105
         Picture         =   "CFGPED04.frx":2D8A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Guarda la Configuracion"
         Top             =   1600
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   37
         Top             =   5460
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
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
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED04.frx":3094
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ayuda en Linea"
         Top             =   950
         Width           =   650
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
         Height          =   650
         Left            =   105
         Picture         =   "CFGPED04.frx":339E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Cerrar y Salir"
         Top             =   300
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -300
         Picture         =   "CFGPED04.frx":34E8
         Top             =   5775
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Pedido de Cliente"
      Height          =   6255
      Left            =   210
      TabIndex        =   10
      Top             =   105
      Width           =   5055
      Begin VB.Frame Frame20 
         BackColor       =   &H00E0E0C0&
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
         Height          =   660
         Left            =   3045
         TabIndex        =   71
         Top             =   2780
         Width           =   1800
         Begin VB.CheckBox Check13 
            BackColor       =   &H00E0E0C0&
            Caption         =   "No Repite Código"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   210
            TabIndex        =   72
            Top             =   210
            Width           =   1560
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Variantes"
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
         Left            =   1785
         TabIndex        =   67
         Top             =   4410
         Width           =   1380
         Begin VB.CheckBox Check12 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Color/Tela"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   68
            Top             =   315
            Width           =   1245
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Listadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   3045
         TabIndex        =   63
         Top             =   1890
         Width           =   1800
         Begin VB.OptionButton Option8 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Moneda de Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   315
            TabIndex        =   65
            Top             =   420
            Width           =   1380
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   64
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00E0E0C0&
         Caption         =   "O/C Cliente"
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
         Left            =   210
         TabIndex        =   61
         Top             =   5250
         Width           =   2145
         Begin VB.CheckBox Check30 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Valida Exist."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1150
            TabIndex        =   75
            Top             =   360
            Width           =   900
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Obliga Informar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   60
            TabIndex        =   62
            Top             =   315
            Width           =   1050
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0C0&
         Caption         =   "O/Fabric."
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
         Left            =   3685
         TabIndex        =   56
         Top             =   5250
         Width           =   1180
         Begin VB.CheckBox Check9 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Generar O/Fab."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   70
            TabIndex        =   57
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   210
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
         Width           =   3750
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Aprobación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1800
         TabIndex        =   42
         Top             =   3465
         Width           =   2235
         Begin VB.OptionButton Option9 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Control Fact. Vencidas"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   1980
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Posterior"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1060
            TabIndex        =   43
            Top             =   260
            Width           =   1140
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Directa"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   115
            TabIndex        =   44
            Top             =   260
            Width           =   1380
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Valorizar"
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
         Left            =   210
         TabIndex        =   32
         Top             =   4410
         Width           =   1545
         Begin VB.CheckBox Check4 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Muestra Importes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   210
            TabIndex        =   33
            Top             =   210
            Width           =   1140
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Ventana"
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
         Left            =   3255
         TabIndex        =   25
         Top             =   4410
         Width           =   1605
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   105
            MaxLength       =   8
            TabIndex        =   26
            ToolTipText     =   "Introduzca el Nombre del Controlador 'PRF'"
            Top             =   315
            Width           =   1460
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0C0&
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
         Height          =   1545
         Left            =   210
         TabIndex        =   23
         Top             =   1890
         Width           =   2745
         Begin VB.OptionButton Option11 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Al Generar"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   170
            TabIndex        =   74
            Top             =   1225
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Al Aprobar"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   73
            Top             =   1225
            Width           =   1245
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Replica por E-mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   66
            Top             =   900
            Width           =   1980
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Impresión Fantasma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   24
            Top             =   680
            Width           =   2115
         End
         Begin VB.CheckBox Check8 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Orden de Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   59
            Top             =   460
            Width           =   2085
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Orden de Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   60
            Top             =   240
            Value           =   1  'Checked
            Width           =   1800
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Crédito C.C."
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
         Left            =   2400
         TabIndex        =   21
         Top             =   5250
         Width           =   1245
         Begin VB.CheckBox Check1 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Verificar Límite "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   70
            TabIndex        =   22
            Top             =   315
            Width           =   1120
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   4095
         TabIndex        =   18
         Top             =   3465
         Width           =   750
         Begin VB.TextBox Text8 
            Height          =   360
            Left            =   140
            TabIndex        =   19
            Text            =   "PC"
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   210
         TabIndex        =   15
         Top             =   3465
         Width           =   1545
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   17
            Top             =   560
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0C0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   16
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4095
         Picture         =   "CFGPED04.frx":540A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   315
         Width           =   750
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   210
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3780
         TabIndex        =   12
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2415
         TabIndex        =   11
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Despacho:"
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
         Left            =   0
         TabIndex        =   55
         Top             =   1260
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   20
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Despacho"
      Height          =   3690
      Left            =   105
      TabIndex        =   0
      Top             =   7980
      Width           =   5160
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFC0&
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
         Height          =   960
         Left            =   2415
         TabIndex        =   29
         Top             =   1365
         Width           =   1545
         Begin VB.CheckBox Check6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Captura"
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
            Left            =   210
            TabIndex        =   31
            ToolTipText     =   "Captura Cola de Impresión de Ordenes de Despacho y/o Pedidos."
            Top             =   600
            Width           =   1170
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   30
            ToolTipText     =   "Suprime Impresión - Enía Datos a Cola de Impresión."
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Generación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2730
         TabIndex        =   27
         Top             =   2415
         Width           =   2175
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Despachar Según Disponibilidad de Stocks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   210
            TabIndex        =   28
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Controles Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   210
         TabIndex        =   7
         Top             =   2415
         Width           =   2325
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   1680
            TabIndex        =   8
            Text            =   " "
            Top             =   550
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Anticipo de Activación de Despacho (dias):"
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
            Left            =   -105
            TabIndex        =   9
            Top             =   315
            Width           =   1665
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "CFGPED04.frx":5714
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
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
         Left            =   210
         TabIndex        =   3
         Top             =   1365
         Width           =   2070
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   5
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   4
            Top             =   580
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
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
         Left            =   4095
         TabIndex        =   1
         Top             =   1365
         Width           =   855
         Begin VB.TextBox Text5 
            Height          =   360
            Left            =   210
            TabIndex        =   2
            Text            =   "OD"
            Top             =   420
            Width           =   450
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario:"
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
      Left            =   210
      TabIndex        =   54
      Top             =   1260
      Width           =   2325
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu tablasubcodigos 
         Caption         =   "Acceso a Tabla de Sub-Códigos"
      End
      Begin VB.Menu cotizaciondetipodecambio 
         Caption         =   "Acces a Cotización de Tipo de Cambio"
      End
      Begin VB.Menu tablatelaycolores 
         Caption         =   "Acceso a Tablas de Telas y Colores"
      End
   End
End
Attribute VB_Name = "CFGPED04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Private Sub Check14_Click()
  If Option6.Value = True Then
    Call MENSERR(24, "Solo para Aprobacion Posterior")
    Check14.Value = 0
End If

End Sub


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Bases de Datos MDB(*.MDB)|*.MDB|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label3 = UCase$(CommonDialog1.FileName)
   '
99 Screen.MousePointer = 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label6 = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command12_Click()
    ACONEC$ = "FLFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   OPX% = ALTERNA%("Colores de Base\Tipos de Tela\Colores de Tela")
   If OPX% = 1 Then
        Call TRALOCAL("TACOLBAS", "")
        VTB% = VENTABU%("TACOLBAS")
        If VTB% >= 0 Then Call TRACENTRAL("TACOLBAS", "")
     ElseIf OPX% = 2 Then
        Call TRALOCAL("TABTELAS", "")
        VTB% = VENTABU%("TABTELAS")
        If VTB% >= 0 Then Call TRACENTRAL("TABTELAS", "")
     ElseIf OPX% = 3 Then
        Call TRALOCAL("TACOLTEL", "")
        VTB% = VENTABU%("TACOLTEL")
        If VTB% >= 0 Then Call TRACENTRAL("TACOLTEL", "")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORDESPA", Text1.TEXT)
    Call GRACONTROL("", "CODDESPA", Text5.TEXT)
    Call GRACONTROL("", "FORPECLI", Text6.TEXT)
    Call GRACONTROL("", "CODPECLI", Text8.TEXT)
    Call GRACONTROL("", "VENCARPE", Text4.TEXT)
    Call GRACONTROL("PEDIDO", "ORIDATOS", Label3)
    Call GRACONTROL("PEDIDO", "FECORTE", Label6)
    '
    MONUME$ = "AUTO"
    If Option1.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de O/D's.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMDESPA", MONUME$)
    '
    MONUME$ = TRIM$(Text3.TEXT)
    Call GRACONTROL$("", "ANTIDESP", MONUME$)
    '
    MONUME$ = "AUTO"
    If Option3.Value = True Then
        Call COMUNI("No Está Disponible la Opción\de Numeración Manual de Pedidos de Clientes.")
        Exit Sub
    End If
    Call GRACONTROL("", "NUMPECLI", MONUME$)
    '
    MONUME$ = "NORMAL": FORCAPED% = 1
    Call GRACONTROL("ADMIPED", "FORMACAR", MONUME$)
    '
    MONUME$ = "NO": VERLICRE% = 0
    If Check1.Value = 1 Then
        MONUME$ = "SI": VERLICRE% = 1
    End If
    Call GRACONTROL("ADMIPED", "VERLICRE", MONUME$)
    '
    MONUME$ = "NO": GENORFAB% = 0
    If Check9.Value = 1 Then
        MONUME$ = "SI": GENORFAB% = 1
    End If
    Call GRACONTROL("ADMIPED", "GENORFAB", MONUME$)
    '
    '\\\OT-06-0512-RG-27/11/06///
    MONUME$ = "NO"
    If Check10.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ADMIPED", "OBLICOM", MONUME$)
    '\\\OT-06-0512-RG-FIN///
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("PEDCLIEN", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO"
    If Check5.Value = 1 Then SUPRINT$ = "SI"
    Call GRACONTROL("ORDESPA", "SUPRINT", SUPRINT$)
    '
    CAPRINT$ = "NO"
    If Check6.Value = 1 Then CAPRINT$ = "SI"
    Call GRACONTROL("CAPRINT", IDENTER$, CAPRINT$)
    '
    CMO$ = ""
    If Check3.Value = 1 Then CMO$ = "REVISTOCK"
    Call GRACONTROL$("ADMIPED", "GENDESPA", CMO$)
    '
    CMO$ = "NO"
    If Check4.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL$("ADMIPED", "MUEPREC", CMO$)
    '\\\OT-05-0236-WAR-02/05/05///
    MONUME$ = "POSTERIOR"
    If Option6.Value = True Then
        MONUME$ = "DIRECT"
      ElseIf Option9.Value = True Then
        MONUME$ = "FACVENCI"
    End If
    '
    Call GRACONTROL("PEDCLI", "MODOAPRO", MONUME$)
    '\\\OT-05-0236-WAR-FIN///
    '\\\OT-06-0439-MH-19/10/06///
    CMO$ = "SI"
    If Check7.Value = 0 Then CMO$ = "NO"
    Call GRACONTROL("PEDIDO", "PRINOVEN", CMO$)
    '
    CMO$ = "NO"
    If Check8.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL("PEDIDO", "PRINODES", CMO$)
    '\\\OT-06-0439-MH-FIN///
    CMO$ = "NO"
    If Check11.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL("ADMIPED", "REPLICAM", CMO$)
    '
    '\\\OT-08-0133-FL-03/04/08///
    CMO$ = "NO"
    If Check12.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL("PEDIDO", "VARCOLOR", CMO$)
    '\\\OT-08-0133-FL-FIN///
    '
    '\\\OT-08-0311-FL-28/05/08///
    CMO$ = "NO"
    If Check13.Value = 1 Then CMO$ = "SI"
    Call GRACONTROL("PEDIDO", "SUPRIREP", CMO$)
    '\\\OT-08-0311-FL-FIN///
    '
    MONUME$ = "PESOS"
    If Option8.Value = True Then MONUME$ = "MONEMORI"
    Call GRACONTROL("ADMIPED", "FORMALIS", MONUME$)
    
    CMO$ = "ALGENERA"
    If Option10.Value = True Then CMO$ = "ALAPROBA"
    Call GRACONTROL("PEDIDO", "IMPRIME", CMO$)
    '
    VALOR$ = "NO"
    If Check30.Value = 1 Then
        VALOR$ = "SI"
    End If
    Call GRACONTROL("", "VALEXIOC", VALOR$)

    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
    Hide
    '
End Sub

Private Sub Command23_Click()
  Call Command2_Click
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") = 2 Then
     Call CARTAMONED
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    '
    Call TRALOCAL("TASUBCOD", "")
10  VTB% = VENTABU("TASUBCOD")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TASUBCOD")
        X$ = REGLEIDO$("!TASUBCOD", U&)
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        CD$ = Left$(X$, 8)
        If Len(CD$) < 1 Then
          Call MENSERR(24, "Sub-Código Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que el CC posea descripcion
        DESC$ = TRIM$(Mid$(X$, 9, 40))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Falta Descripción / Referencia")
          GoTo 10
        End If
        ' Valida que el CC no este repetido
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TASUBCOD", U2&)
          If TRIM$(Mid$(X$, 1, 8)) = TRIM$(Mid$(Y$, 1, 8)) Then
            Call MENSERR(24, "Sub-Código Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TASUBCOD", "")
    End If
    '
99  Command9.Enabled = True
    '
End Sub

Private Sub cotizaciondetipodecambio_Click()
   Call Command8_Click
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Frame9.BackColor = &H8000000F
      Me.Frame10.BackColor = &H8000000F
      Me.Frame11.BackColor = &H8000000F
      Me.Frame12.BackColor = &H8000000F
      Me.Frame13.BackColor = &H8000000F
      Me.Frame14.BackColor = &H8000000F
      Me.Frame15.BackColor = &H8000000F
      Me.Frame16.BackColor = &H8000000F
      Me.Frame17.BackColor = &H8000000F
      Me.Frame18.BackColor = &H8000000F
      Me.Frame19.BackColor = &H8000000F
      Me.Frame20.BackColor = &H8000000F
      '
      Me.Option1.BackColor = &H8000000F
      Me.Option10.BackColor = &H8000000F
      Me.Option11.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      Me.Option8.BackColor = &H8000000F
      Me.Option9.BackColor = &H8000000F
      '
      Me.Check1.BackColor = &H8000000F
      Me.Check10.BackColor = &H8000000F
      Me.Check11.BackColor = &H8000000F
      Me.Check12.BackColor = &H8000000F
      Me.Check13.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check30.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      Me.Check6.BackColor = &H8000000F
      Me.Check7.BackColor = &H8000000F
      Me.Check8.BackColor = &H8000000F
      Me.Check9.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
    End If

    EPAC$ = TRIM$(UCase$(EMPREAC$))
    EPLI$ = TRIM$(UCase$(EMPRELI$))
    If InStr(EPAC$, "ZETRONIC") < 1 And InStr(EPLI$, "ZETRONIC") < 1 Then
      Frame15.Visible = False
      Height = Picture2.Height + 700
    End If
    '
    Text1.TEXT = CONTROL$("", "FORDESPA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text6.TEXT = CONTROL$("", "FORPECLI")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
    VENCARPE$ = CONTROL$("", "VENCARPE")
    If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
    Text4.TEXT = VENCARPE$
    '
    '\\\OT-06-0512-RG-27/11/06///
    If CONTROL$("ADMIPED", "OBLICOM") = "SI" Then Check10.Value = 1
    '\\\OT-06-0512-RG-FIN///
    If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then Check1.Value = 1
    If CONTROL$("ADMIPED", "GENORFAB") = "SI" Then Check9.Value = 1
    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
       Check5.Value = 1
    End If
    '
    If IDENTER$ <> "" Then
      If CONTROL$("CAPRINT", IDENTER$) = "SI" Then
        Check6.Value = 1
      End If
    End If
    '\\\OT-05-0236-WAR-02/05/05///
     If CONTROL("PEDCLI", "MODOAPRO") = "DIRECT" Then
          Option6.Value = True
       ElseIf CONTROL("PEDCLI", "MODOAPRO") = "POSTERIOR" Then
          Option5.Value = True
       ElseIf CONTROL("PEDCLI", "MODOAPRO") = "FACVENCI" Then
          Option9.Value = True
     End If
     
     '\\\OT-05-0236-WAR-FIN///
    If CONTROL$("", "NUMDESPA") = "MANUAL" Then Option1.Value = True
    If CONTROL$("", "NUMPECLI") = "MANUAL" Then Option3.Value = True
    If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then Option5.Value = True
    If CONTROL$("ADMIPED", "FORMALIS") = "MONEMORI" Then Option8.Value = True
    
    Text3.TEXT = TRIM$(CONTROL$("", "ANTIDESP"))
    If XVALO(Text3.TEXT) < 1 Then Text3.TEXT = "3"
    '
    CMO$ = TRIM$(CONTROL$("", "CODDESPA")): If CMO$ <> "" Then Text5.TEXT = CMO$
    CMO$ = TRIM$(CONTROL$("", "CODPECLI")): If CMO$ <> "" Then Text8.TEXT = CMO$
    '
    If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then Check3.Value = 1
    If CONTROL$("ADMIPED", "MUEPREC") = "SI" Then Check4.Value = 1
    '
    Label3 = CONTROL$("PEDIDO", "ORIDATOS")
    Label6 = CONTROL$("PEDIDO", "FECORTE")
    '\\\OT-06-0439-MH-19/10/06///
    Check7.Value = 1
    If CONTROL$("PEDIDO", "PRINOVEN") = "NO" Then Check7.Value = 0
    Check8.Value = 0
    If CONTROL$("PEDIDO", "PRINODES") = "SI" Then Check8.Value = 1
    '\\\OT-06-0439-MH-FIN///
    If CONTROL$("ADMIPED", "REPLICAM") = "SI" Then Check11.Value = 1
    '
    '\\\OT-08-0133-FL-03/04/08///
    If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
       Check12.Value = 1
       Command13.Visible = True
    End If
    '\\\OT-08-0133-FL-FIN///
    '
    '\\\OT-08-0311-FL-28/05/08///
    If CONTROL$("PEDIDO", "SUPRIREP") = "SI" Then Check13.Value = 1
    '\\\OT-08-0311-FL-FIN///
    '
    Option11.Value = True
    If CONTROL$("PEDIDO", "IMPRIME") = "ALAPROBA" Then Option10.Value = True
    '
    Check30.Value = 0
    If CONTROL$("", "VALEXIOC") = "SI" Then Check30.Value = 1
    '
End Sub






Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub tablasubcodigos_Click()
   Call Command9_Click
End Sub

Private Sub tablatelaycolores_Click()
   Call Command13_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub


Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
End Sub


