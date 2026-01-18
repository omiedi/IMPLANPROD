VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BkpWaldb 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Rutas de Base Waldbott  de Origen y Lectura"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Cambiar Estado de Bloqueo"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   7080
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   4200
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   8655
      Left            =   10380
      ScaleHeight     =   8595
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Width           =   1170
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
         Height          =   500
         Left            =   120
         Picture         =   "BkpWaldb.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7560
         Width           =   645
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "BkpWaldb.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "BkpWaldb.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8160
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   8895
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame2 
         Caption         =   "Còdigo Rango de Liner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   45
         Top             =   7440
         Width           =   9840
         Begin VB.CheckBox Check21 
            Caption         =   "Validar Cierre Mayor al Proyectado-Tolerancia: %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            TabIndex        =   52
            Top             =   810
            Width           =   4350
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8355
            TabIndex        =   51
            ToolTipText     =   "PORCENTAJE DE TOLERANCIA A CERRAR SUPERANDO LA CANTIDAD PROYECTADA"
            Top             =   840
            Width           =   570
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   " "
            Top             =   615
            Width           =   3750
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
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
            Left            =   3660
            TabIndex        =   47
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   " "
            ToolTipText     =   "Doble Clic para borrar"
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"BkpWaldb.frx":2376
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   15
            Left            =   3960
            TabIndex        =   50
            Top             =   120
            Width           =   5805
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Rango : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   -120
            TabIndex        =   49
            Top             =   360
            Width           =   2325
         End
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Text            =   "BkpWaldb.frx":2408
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   5640
         Width           =   4680
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Text            =   "BkpWaldb.frx":240A
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   4800
         Width           =   4680
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Text            =   "BkpWaldb.frx":240C
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   5640
         Width           =   4680
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "BkpWaldb.frx":240E
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   4800
         Width           =   4680
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Text            =   "BkpWaldb.frx":2410
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   3960
         Width           =   4680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Etiqueta de Impresión Prod.Terminado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   29
         Top             =   6360
         Width           =   4680
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   33
            Text            =   " "
            Top             =   240
            Width           =   1380
         End
         Begin VB.CommandButton Command10 
            Caption         =   "."
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
            Left            =   3660
            TabIndex        =   32
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   " "
            Top             =   615
            Width           =   3750
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   680
            Left            =   3960
            Picture         =   "BkpWaldb.frx":2412
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   -120
            TabIndex        =   34
            Top             =   360
            Width           =   2325
         End
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "BkpWaldb.frx":271C
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   3120
         Width           =   4680
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "BkpWaldb.frx":271E
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   2280
         Width           =   4680
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "BkpWaldb.frx":2720
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   600
         Width           =   4680
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   5280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "BkpWaldb.frx":2722
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   1440
         Width           =   4680
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "BkpWaldb.frx":2724
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   3960
         Width           =   4680
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "BkpWaldb.frx":2726
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   3120
         Width           =   4680
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "BkpWaldb.frx":2728
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   2280
         Width           =   4680
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   6360
         Width           =   3100
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "BkpWaldb.frx":272A
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   1440
         Width           =   4680
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "BkpWaldb.frx":272C
         ToolTipText     =   "Doble Clic para seleccionar Ruta"
         Top             =   600
         Width           =   4680
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de INTERREMI2 Movisto P/Conteplast"
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
         Index           =   14
         Left            =   0
         TabIndex        =   44
         Top             =   5400
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de INTERREMI Movisto P/Conteplast"
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
         Index           =   13
         Left            =   0
         TabIndex        =   42
         Top             =   4560
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión de O.F.Cortes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   5040
         TabIndex        =   40
         Top             =   5400
         Width           =   4965
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión de O.F. Bolsones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   5040
         TabIndex        =   38
         Top             =   4560
         Width           =   4965
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión de O.F. Bolsas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   5040
         TabIndex        =   36
         Top             =   3720
         Width           =   4965
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión Liners"
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
         Index           =   9
         Left            =   5040
         TabIndex        =   28
         Top             =   2880
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión Orden de Fabricación de Cintillas"
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
         Index           =   8
         Left            =   5040
         TabIndex        =   26
         Top             =   2040
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel Impresión Planilla de Venta   "
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
         Index           =   6
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Excel Impresión Cotización Planilla de Venta   "
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
         Index           =   7
         Left            =   5040
         TabIndex        =   23
         Top             =   1200
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel Impresión Doblado"
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
         Index           =   5
         Left            =   0
         TabIndex        =   20
         Top             =   3720
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel Impresión Etiqueta"
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
         Index           =   4
         Left            =   0
         TabIndex        =   18
         Top             =   2880
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel Impresión Ficha Técnica"
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
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   2040
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Impresion Cotización "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   5160
         TabIndex        =   14
         Top             =   6270
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Waldbott de Lectura"
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
         Index           =   0
         Left            =   -480
         TabIndex        =   11
         Top             =   1200
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Waldbott Dropbox"
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
         Index           =   3
         Left            =   -480
         TabIndex        =   9
         Top             =   360
         Width           =   5085
      End
   End
End
Attribute VB_Name = "BkpWaldb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check21_Click()
If Check21.Value = 1 Then 'CUANDO SE DESMARCA LO MUESTRA VACIO, SI LO MARCA TRAE EL VALOR QUE TENIA
    Text18 = XVALO(CONTROL("ORDEFABR", "TOCIEPRO"))
Else
    Text18 = "" 'SI GRABA DESMARCADO QUEDA EN 0
End If
End Sub
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command10_Click()
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text12.SetFocus
      Else
        Text12.TEXT = CODFOR$
    End If

End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("COTICONT", "ARCDROPB", Text1.TEXT)
    Call GRACONTROL("CONWALDB", "RUTAWALD", Text2.TEXT)
    Call GRACONTROL("COTICONT", "EXCEFTEC", Text3.TEXT)
    Call GRACONTROL("COTICONT", "EXCEETIQ", Text4.TEXT)
    Call GRACONTROL("COTICONT", "EXCEDOBL", Text5.TEXT)
    Call GRACONTROL("COTICONT", "EXCEPVTA", Text6.TEXT) 'PLANILLA DE VENTA
    Call GRACONTROL("COTICONT", "EXCCOTPV", Text7.TEXT) 'COTIZACION PLANILLA DE VENTA
    Call GRACONTROL("COTICONT", "EXFABCIN", Text8.TEXT) 'ORDEN DE FABRICACION DE CINTILLA
    Call GRACONTROL("COTICONT", "EXCLINER", Text9.TEXT) 'LINER
    Call GRACONTROL("ETIQPRTE", "ETPROTER", Text12.TEXT) 'ETIQUETA PRODUCOTO TERMINADO
    Call GRACONTROL("PLANVENT", "IMPOFBOL", Text10.TEXT) 'PLANTILLA EXCEL IMPRE O.F. BOLSAS
    Call GRACONTROL("PLANVENT", "IMPOFBIG", Text11.TEXT) 'PLANTILLA EXCEL IMPRE O.F. BOLSONES BIG BAG
    Call GRACONTROL("ORDEFABR", "IMPOFCOR", Text13.TEXT) 'PLANTILLA EXCEL IMPRE O.F. CORTES
    Call GRACONTROL("CONWALDB", "ESCRMOVI", Text14.TEXT) 'RUTA DONDE ESCRIBE EN MDB LOS REGISTROS DUPLICADOS CUANDO REALIZA EL REMITO
    Call GRACONTROL("CONWALDB", "ESCRMOV2", Text15.TEXT) 'RUTA DONDE ESCRIBE (interemi2)EN MDB LOS REGISTROS DUPLICADOS CUANDO REALIZA EL REMITO
    Call GRACONTROL("CONFLINE", "CORANLIN", Text16.TEXT) 'CODIGO RANGO DE LINER PARA AGREEGAR A LA ESTRCUTURA SI LINER = SI
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    
    Call GRACONTROL(IDENTER$, "PRINTSER", Combo2.TEXT)
    Call GRACONTROL("*", "PRINTSER", Combo2.TEXT)

    MONUME$ = "NO"
    If Check21.Value = 1 Then
        MONUME$ = "SI"
    End If
    Call GRACONTROL("ORDEFABR", "VACIECAN", MONUME$) 'VALIDAR AL CIERRE DE LA OF QUE LA CANTIDAD NO SUPERE LA PROYECTADA
    '
    Text18.TEXT = XVALO(Text18.TEXT)
    Call GRACONTROL("ORDEFABR", "TOCIEPRO", Text18.TEXT) 'PORCENTAJE DE TOLERANCIA A CERRAR SUPERANDO LA CANTIDAD PROYECTADA

    Call FINAL("")
    '
End Sub



Private Sub Command3_Click()
    If CONTROL("BASEWALD", "ESTAGRAB") = "SI" Then
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
    Else
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "SI")
    End If
    Command3.Caption = "Cambiar Estado de Bloqueo (" & CONTROL("BASEWALD", "ESTAGRAB") & ")"
End Sub

Private Sub Command4_Click()
    Erase RESP_ZELE_VECT
    '29/09/2020
    Campos$ = "Cód.Rango/A10;" '1
    Campos$ = Campos$ + "Descripción/A64;"      '2
    Campos$ = Campos$ + "Cód.M.P./A16;"           '3
    Campos$ = Campos$ + "DESCRIPCION/A28;"      '4
    Campos$ = Campos$ + "Costo/F10.4;"      '5
    Campos$ = Campos$ + "CIMAT/A0"      '6
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    SQLX$ = "SELECT * FROM CODIGORANGO "
    SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "C.CODRANGO ,C.Descripcion ,C.CODMATPRIMA ,C.DESCRIMATPRIMA,M.COSUNI,C.CIMATPRIMA"
    TABLA$ = "CODIGORANGO C WITH(NOLOCK) INNER JOIN MASTER M ON C.CIMATPRIMA = M.CODINT"
    CONDI$ = "C.CODIGORANGO_ID > 0 Order by C.CODRANGO"
    '
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1
    Text16 = RESP_ZELE_VECT(1)
    Text17 = RESP_ZELE_VECT(2)
End Sub

Private Sub Command61_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text12.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL("COTICONT", "ARCDROPB")
    Text2.TEXT = CONTROL("CONWALDB", "RUTAWALD")
    Text3.TEXT = CONTROL("COTICONT", "EXCEFTEC")
    Text4.TEXT = CONTROL("COTICONT", "EXCEETIQ")
    Text5.TEXT = CONTROL("COTICONT", "EXCEDOBL")
    Text6.TEXT = CONTROL("COTICONT", "EXCEPVTA")
    Text7.TEXT = CONTROL("COTICONT", "EXCCOTPV")
    Text8.TEXT = CONTROL("COTICONT", "EXFABCIN") 'ORDEN DE FABRICACION DE CINTILLA
    Text9.TEXT = CONTROL("COTICONT", "EXCLINER")
    Text12.TEXT = CONTROL("ETIQPRTE", "ETPROTER") 'ETIQUETA PRODUCOTO TERMINADO
    Text10.TEXT = CONTROL("PLANVENT", "IMPOFBOL")
    Text11.TEXT = CONTROL("PLANVENT", "IMPOFBIG") 'IMPRESION DE BOLSONES BIG BAGS
    Text13.TEXT = CONTROL("ORDEFABR", "IMPOFCOR")
    Text14.TEXT = CONTROL("CONWALDB", "ESCRMOVI")
    Text15.TEXT = CONTROL("CONWALDB", "ESCRMOV2") 'rtua interemi2
    Text16.TEXT = CONTROL("CONFLINE", "CORANLIN") ''CODIGO RANGO DE LINER PARA AGREEGAR A LA ESTRCUTURA SI LINER = SI
    Command3.Caption = "Cambiar Estado de Bloqueo (" & CONTROL("BASEWALD", "ESTAGRAB") & ")"
    '
    Text18 = XVALO(CONTROL("ORDEFABR", "TOCIEPRO")) 'PORCENTAJE DE TOLERANCIA A CERRAR SUPERANDO LA CANTIDAD PROYECTADA
    If CONTROL("ORDEFABR", "VACIECAN") = "SI" Then Check21.Value = 1  'VALIDAR AL CIERRE DE LA OF QUE LA CANTIDAD NO SUPERE LA PROYECTADA
    
    PDES$ = CONTROL$(IDENTER$, "PRINTSER")
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRINTSER")
    Combo2.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo2.TEXT = PDES$
      End If
    Next
    '
End Sub




Private Sub Text1_DblClick()
     SELCARP07.Caption = "Seleccione Carpeta "
     SELCARP07.Show 1
     Text1 = SELCARP07.Label1
End Sub

Private Sub Text10_DblClick()
   Text10 = SELECCION_Archivo$("*", CMD1)
End Sub
Private Sub Text11_DblClick()
   Text11 = SELECCION_Archivo$("*", CMD1)
End Sub
Private Sub Text12_Change()
    CODFOR$ = TRIM$(Text12)
    Text22 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub Text13_DblClick()
   Text13 = SELECCION_Archivo$("*", CMD1)
End Sub

Private Sub Text14_DblClick()
   Text14 = SELECCION_Archivo$("*", CMD1)

End Sub

Private Sub Text16_Change()
      Text17 = DESCRIPCIONCODIGORANGO(TRIM$(Text16))
End Sub

Private Sub Text16_DblClick()
   Text16 = ""
   Text17 = ""
End Sub


Private Sub Text2_DblClick()
   Text2 = SELECCION_Archivo$("MDB", CMD1)
End Sub

Private Sub Text3_DblClick()
   Text3 = SELECCION_Archivo$("*", CMD1)
End Sub

Private Sub Text4_DblClick()
   Text4 = SELECCION_Archivo$("*", CMD1)
End Sub

Private Sub Text5_DblClick()
   Text5 = SELECCION_Archivo$("*", CMD1)
End Sub
Private Sub Text6_DblClick()
   Text6 = SELECCION_Archivo$("*", CMD1)
End Sub
Private Sub Text7_DblClick()
   Text7 = SELECCION_Archivo$("*", CMD1)
End Sub

Private Sub Text8_DblClick()
   Text8 = SELECCION_Archivo$("*", CMD1)
End Sub
Private Sub Text9_DblClick()
   Text9 = SELECCION_Archivo$("*", CMD1)
End Sub
