VERSION 5.00
Begin VB.Form FORCOMPRA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FORCOMPRA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   4350
         Left            =   50
         TabIndex        =   87
         Top             =   1580
         Visible         =   0   'False
         Width           =   10830
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   105
            ScaleHeight     =   375
            ScaleWidth      =   10515
            TabIndex        =   93
            Top             =   0
            Width           =   10580
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción General - Especificaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   0
               TabIndex        =   94
               Top             =   105
               Width           =   10515
            End
         End
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Left            =   1470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   430
            Width           =   7890
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00C0C0A0&
            Caption         =   "Artículo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   700
            Left            =   105
            TabIndex        =   88
            Top             =   3570
            Width           =   10575
            Begin VB.CommandButton Command18 
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
               Left            =   2520
               TabIndex        =   90
               Top             =   280
               Width           =   180
            End
            Begin VB.TextBox Text18 
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
               Left            =   210
               TabIndex        =   89
               Text            =   " "
               Top             =   280
               Width           =   2325
            End
            Begin VB.Label Label8 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Left            =   2835
               TabIndex        =   91
               Top             =   285
               Width           =   7515
            End
         End
         Begin VB.Line Line6 
            X1              =   10680
            X2              =   10680
            Y1              =   0
            Y2              =   3465
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   3060
            Left            =   105
            TabIndex        =   96
            Top             =   435
            Width           =   1375
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   3060
            Left            =   9345
            TabIndex        =   95
            Top             =   435
            Width           =   1320
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Texto Libre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   5985
         TabIndex        =   100
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Pos.IVA - CUIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1065
         Left            =   5670
         TabIndex        =   97
         Top             =   15
         Width           =   1905
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   105
            TabIndex        =   99
            TabStop         =   0   'False
            Text            =   " "
            Top             =   630
            Width           =   1695
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   105
            TabIndex        =   98
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   82
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   1115
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   2250
         Left            =   10450
         TabIndex        =   81
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Pendientes de Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1960
         Left            =   130
         TabIndex        =   44
         Top             =   5985
         Width           =   10580
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            Height          =   1730
            Left            =   9345
            TabIndex        =   80
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   1730
            Left            =   8715
            TabIndex        =   79
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   1730
            Left            =   7665
            TabIndex        =   78
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1730
            Left            =   5985
            TabIndex        =   77
            Top             =   210
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   1730
            Left            =   1890
            TabIndex        =   76
            Top             =   210
            Width           =   15
         End
         Begin VB.ListBox List2 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   0
            MultiSelect     =   1  'Simple
            TabIndex        =   47
            Top             =   750
            Width           =   10565
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   45
            Top             =   315
            Width           =   10565
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Código          Descripción                           Cantid.   F.Entreg  Atr    O/C"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   30
               TabIndex        =   46
               Top             =   105
               Width           =   10515
            End
         End
         Begin VB.Line Line2 
            X1              =   10565
            X2              =   10565
            Y1              =   315
            Y2              =   1940
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   10560
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   7
         Left            =   6250
         TabIndex        =   43
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   6
         Left            =   8665
         TabIndex        =   42
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   5
         Left            =   7465
         TabIndex        =   41
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   4
         Left            =   9400
         TabIndex        =   40
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   3
         Left            =   5725
         TabIndex        =   39
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   2
         Left            =   4675
         TabIndex        =   38
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2250
         Index           =   1
         Left            =   4100
         TabIndex        =   37
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2280
         Index           =   0
         Left            =   2050
         TabIndex        =   36
         Top             =   1575
         Width           =   15
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Valorizar en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7665
         TabIndex        =   32
         Top             =   840
         Width           =   3060
         Begin VB.TextBox Text16 
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
            Height          =   240
            Left            =   900
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
            Width           =   2010
         End
         Begin VB.TextBox Text15 
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
            Height          =   240
            Left            =   210
            MaxLength       =   2
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   " "
            Top             =   280
            Width           =   435
         End
         Begin VB.CommandButton Command12 
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
            Height          =   240
            Left            =   630
            TabIndex        =   33
            Top             =   280
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0A0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   800
         Left            =   7665
         TabIndex        =   29
         Top             =   15
         Width           =   1380
         Begin VB.CommandButton Command7 
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
            Left            =   1050
            TabIndex        =   31
            Top             =   420
            Width           =   195
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Left            =   105
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.TextBox Text6 
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   26
         Text            =   " "
         Top             =   1100
         Width           =   4065
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   160
         TabIndex        =   23
         Top             =   1995
         Width           =   10565
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   120
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6950
         Left            =   10920
         TabIndex        =   17
         Top             =   105
         Width           =   850
         Begin VB.CommandButton Command22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Precios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Actualizar Precios de Proveedores"
            Top             =   6170
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":06D6
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   1542
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":0B18
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   3084
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proved."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":0E22
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   2313
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":112C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   3855
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":1276
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   771
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":1580
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Consultas y Listados"
            Top             =   4627
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORCOMPRA.frx":188A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   5395
            Width           =   855
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   760
         Left            =   10920
         Picture         =   "FORCOMPRA.frx":1B94
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
         Top             =   7200
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   160
         ScaleHeight     =   375
         ScaleWidth      =   10500
         TabIndex        =   15
         Top             =   1575
         Width           =   10565
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción      U/S     Cant  U/C      Cant   Pr.Unit % Dc  F.Entreg"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   30
            TabIndex        =   16
            ToolTipText     =   "Haga Doble-Click para Blanquear Pre-Carga y Habilitar Carga Manual"
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
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
         Left            =   9615
         TabIndex        =   8
         Text            =   " "
         Top             =   105
         Width           =   1120
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   9615
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   960
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10560
         TabIndex        =   7
         Top             =   525
         Width           =   195
      End
      Begin VB.TextBox Text5 
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
         Height          =   285
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
         Width           =   3225
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   440
         Width           =   4065
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4065
      End
      Begin VB.CommandButton Command1 
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
         Left            =   920
         TabIndex        =   1
         Top             =   1115
         Width           =   195
      End
      Begin VB.TextBox Text1 
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
         Left            =   160
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   1100
         Width           =   750
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2050
         Left            =   0
         TabIndex        =   48
         Top             =   3960
         Width           =   10800
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0A0&
            Caption         =   "Resumen Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   1965
            Left            =   6615
            TabIndex        =   49
            Top             =   0
            Width           =   4130
            Begin VB.TextBox Text17 
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
               Left            =   1365
               MaxLength       =   14
               TabIndex        =   54
               Top             =   1260
               Width           =   750
            End
            Begin VB.TextBox Total 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   2415
               TabIndex        =   84
               Top             =   1260
               Width           =   1500
            End
            Begin VB.TextBox PORDESCU 
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
               Left            =   735
               MaxLength       =   14
               TabIndex        =   52
               Top             =   630
               Width           =   1380
            End
            Begin VB.TextBox Total 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   2415
               TabIndex        =   56
               TabStop         =   0   'False
               Text            =   " "
               Top             =   630
               Width           =   1500
            End
            Begin VB.TextBox Total 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   2415
               TabIndex        =   55
               Top             =   315
               Width           =   1500
            End
            Begin VB.TextBox Text12 
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
               Left            =   1365
               MaxLength       =   14
               TabIndex        =   53
               Top             =   945
               Width           =   750
            End
            Begin VB.TextBox Total 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   2415
               TabIndex        =   51
               Top             =   945
               Width           =   1500
            End
            Begin VB.TextBox Total 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   2415
               TabIndex        =   50
               Top             =   1575
               Width           =   1500
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
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
               Left            =   2205
               TabIndex        =   86
               Top             =   1320
               Width           =   225
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Percepcion:"
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
               Index           =   3
               Left            =   -1680
               TabIndex        =   85
               Top             =   1320
               Width           =   2955
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
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
               Left            =   2205
               TabIndex        =   63
               Top             =   720
               Width           =   225
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Desc:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -1260
               TabIndex        =   62
               Top             =   675
               Width           =   1905
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal:"
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
               Index           =   0
               Left            =   1365
               TabIndex        =   61
               Top             =   375
               Width           =   960
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "I.V.A.:"
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
               Index           =   2
               Left            =   -630
               TabIndex        =   60
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Orden de Compra:"
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
               Index           =   4
               Left            =   -630
               TabIndex        =   59
               ToolTipText     =   "Haga un Doble-Click para Emitir CON / SIN Precios"
               Top             =   1620
               Width           =   2955
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
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
               Left            =   2205
               TabIndex        =   58
               Top             =   1020
               Width           =   225
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   105
               TabIndex        =   57
               ToolTipText     =   "Doble-Click para 'Dólares'"
               Top             =   260
               Width           =   1275
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0A0&
            Caption         =   "Condiciones Comerciales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   1965
            Left            =   150
            TabIndex        =   64
            Top             =   0
            Width           =   6270
            Begin VB.TextBox Text10 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   1155
               TabIndex        =   69
               Top             =   890
               Width           =   4950
            End
            Begin VB.CommandButton Command21 
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
               Height          =   250
               Index           =   0
               Left            =   1575
               TabIndex        =   72
               Top             =   315
               Width           =   180
            End
            Begin VB.TextBox TEXT21 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   0
               Left            =   1155
               TabIndex        =   71
               Text            =   " "
               Top             =   315
               Width           =   435
            End
            Begin VB.TextBox Eco21 
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
               Height          =   250
               Index           =   0
               Left            =   1890
               TabIndex        =   70
               Top             =   315
               Width           =   4215
            End
            Begin VB.CommandButton Command21 
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
               Height          =   250
               Index           =   1
               Left            =   1575
               TabIndex        =   68
               Top             =   600
               Width           =   180
            End
            Begin VB.TextBox TEXT21 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   1
               Left            =   1155
               TabIndex        =   67
               Text            =   " "
               Top             =   600
               Width           =   435
            End
            Begin VB.TextBox Eco21 
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
               Height          =   250
               Index           =   1
               Left            =   1890
               TabIndex        =   66
               Text            =   " "
               Top             =   600
               Width           =   4215
            End
            Begin VB.TextBox Text13 
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
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   65
               Top             =   1260
               Width           =   4950
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L.Entrega:"
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
               Index           =   2
               Left            =   -3045
               TabIndex        =   83
               Top             =   930
               Width           =   4110
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "C.Pago:"
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
               Index           =   0
               Left            =   -525
               TabIndex        =   75
               Top             =   370
               Width           =   1590
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Embalaje:"
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
               Index           =   1
               Left            =   -3045
               TabIndex        =   74
               Top             =   655
               Width           =   4110
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Observac:"
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
               Index           =   3
               Left            =   -3045
               TabIndex        =   73
               Top             =   1260
               Width           =   4110
            End
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5040
         TabIndex        =   103
         Top             =   1100
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PQ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   4725
         TabIndex        =   102
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   101
         Top             =   1470
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Line Line5 
         X1              =   180
         X2              =   10720
         Y1              =   3850
         Y2              =   3850
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   750
         Left            =   165
         TabIndex        =   24
         Top             =   -30
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   105
         Width           =   960
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1575
         Y2              =   3860
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   8400
         TabIndex        =   14
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   8820
         TabIndex        =   13
         Top             =   550
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   105
         TabIndex        =   12
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archima 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Master 
         Caption         =   "Articulos de Stock"
      End
      Begin VB.Menu Provedor 
         Caption         =   "Proveedores y Contratistas"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig 
         Caption         =   "Números Originales"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Orcom 
      Caption         =   "Ordenes de Compra"
      Begin VB.Menu Nue_Orco 
         Caption         =   "Nueva Orden de Compra"
         Begin VB.Menu Directa 
            Caption         =   "O/C Directa"
         End
         Begin VB.Menu Por_MRP 
            Caption         =   "O/C por MRP"
         End
      End
      Begin VB.Menu Modi_Ocom 
         Caption         =   "Modificar O/Compra"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu Ver_Ocom 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu MarItCum 
         Caption         =   "Marcar Items Cumplidos"
      End
      Begin VB.Menu plh22 
         Caption         =   "-"
      End
      Begin VB.Menu Anul_Ocom 
         Caption         =   "Anular O/Compra"
      End
      Begin VB.Menu Desanuoc 
         Caption         =   "Revertir Anulación"
      End
      Begin VB.Menu plh6 
         Caption         =   "-"
      End
      Begin VB.Menu Maconsi1 
         Caption         =   "Materiales Consignados"
         Begin VB.Menu Entrecon 
            Caption         =   "Entregar Materiales"
         End
         Begin VB.Menu DevoConsi 
            Caption         =   "Devoluciones"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu CoCoPen 
         Caption         =   "Control de Compras Pendientes"
         Begin VB.Menu ConuOC 
            Caption         =   "Por Numero de O/C"
            Begin VB.Menu Screen1 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print1 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoProv 
            Caption         =   "Por Proveedor"
            Begin VB.Menu Screen2 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print2 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu CoMat 
            Caption         =   "Por Material"
            Begin VB.Menu Screen3 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print3 
               Caption         =   "Imprimir / Exportar"
            End
         End
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu Maconsi 
            Caption         =   "Material en Consignación"
            Begin VB.Menu Screen4 
               Caption         =   "Por Pantalla"
            End
            Begin VB.Menu Print4 
               Caption         =   "Imprimir / Exportar"
            End
         End
      End
      Begin VB.Menu CoUlCom 
         Caption         =   "Control Ultimas Compras"
         Begin VB.Menu UcoPro 
            Caption         =   "Control por Proveedor"
         End
         Begin VB.Menu UcoMat 
            Caption         =   "Control por Material"
         End
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuracion "
      Begin VB.Menu Setup 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Nuorig1 
         Caption         =   "Números Originales"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FORCOMPRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim MODOULPRE%

Private Sub Anul_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOMP.Command14_Click
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
        FORIPRE$ = CONTROL$("", "FORMOCOM")
        If FORIPRE$ <> "" Then
          If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
            AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
            ACOLDE = Val(Mid$(AA1$, 2))
            If ACOLDE > 0 Then
              If ACOLDE <= 128 Then
                TEPRUE.Font = Text28.Font
                TEPRUE.FontSize = Text28.FontSize
                TEPRUE.FontBold = Text28.FontBold
                TEPRUE = String$(ACOLDE + 3, "*")
                ATEX = TEPRUE.Width + 50
                '
                ANCHOTO = Label14.Width + Text28.Width + Label15.Width
                Label14.Width = (ANCHOTO - ATEX) / 2
                Label15.Width = (ANCHOTO - ATEX) / 2
                Text28.Left = Label15.Left + Label15.Width
                Text28.Width = ATEX
                Label14.Left = Text28.Left + Text28.Width
                '
              End If
            End If
          End If
        End If
        '
        Frame16.Top = 1570
        Frame16.Visible = True
        Frame2.Top = 5985
        Frame9.Visible = False
        DoEvents
      Else
        Frame16.Visible = False
        Frame2.Top = 3960
        Frame9.Visible = True
   End If
End Sub

Private Sub Command1_Click()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       If Frame16.Visible = True Then
         Text1.TEXT = TRIM$(Str$(NCLIE%))
         Call CARDACLI
         Text28 = ""
         Text18 = ""
         Exit Sub
       End If
       '
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command10_Click()
   Call SELECHO("PROVEPLA")
   NCLIE% = Val(TRIM$(VALACT1$("PROVEPLA")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARCOANT
       If ULTREG&("!OCOMDETA") > 0 Then
            Call CARLISCO
            Call CALTOCOM
         Else
            Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
3   Call SELECHO("TAMONED")
    MMI% = Val(VALACT1$("TAMONED"))
    If MMI% > 0 Then
      Text15 = TRIM$(Str$(MMI%))
    End If
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   OPCICOMP.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call GRACOMPRA(OKX%)
   If OKX% > 0 Then
     Call GRAOCSIM
     '\\\OT-06-0005-WAR-10/01/06///
10   SALIOCOM.Show 1
     Call FRESHALL
     If SALIOCOM.Check1.Value < 1 Then
       If SALIOCOM.Check2.Value < 1 Then
         Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
         GoTo 10
       End If
     End If
     If SALIOCOM.Check1.Value = 1 Then
        Call PRICOMPRA   ' A IMPRESORA
     End If
     If SALIOCOM.Check2.Value = 1 Then
15        SENDBYMAIL$ = ""
        PORMAIL% = 1
        NCXX% = (-1) * Val(Text1)
        EMAI$ = EMAILCLI$(NCXX%)
        If EMAI$ = "" Or InStr(EMAI$, "@") < 2 Then
          Call MENSERR(24, "Imposible Enviar por E-Mail.\   \Dirección Electrónica Inexistente\o no Válida.")
          Call PRICOMPRA
          GoTo 25
        End If
        SENDBYMAIL$ = EMAI$
        '
        For Each prx1 In Printers
           If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then GoTo 20
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        Call PRICOMPRA   ' A IMPRESORA
        GoTo 25
        '
20        Call PRICOMPRA
        Call RECONECTA
        Exit Sub
     End If
     '\\\OT-06-0005-WAR-FIN///
     '
25   Call BLANFORMU
     Call SETULTREG("!OCOMDETA", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
   Command14.Enabled = True
End Sub


Private Sub Command18_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call SELECHO("MASTER")
   Text18 = VALACT1$("MASTER")
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    If DERACCE%("ACOSPROV", "Actualizacion de Costos por Proveedore") >= 2 Then
      HOII% = HOY(HOS$)
3     VDEF$ = String$(6, 0) + MKI$(2) + MKI$(HOII%) + MKI$(HOII%)
5     OBX$ = OBJPLA$("AJUCOSPRO", VDEF$)
      If OBX$ = "" Then GoTo 99
      '
10    VDEF$ = OBX$
      NCX1% = CVI(Left$(OBX$, 2))
      If NCX1% < 1 Then GoTo 5
      If ECOARCH$("ACREDOR", MKI$(NCX1%)) = "" Then
        Call MENSERR(24, "Numero de Proveedor Inexistente o no Válido")
        GoTo 5
      End If
      '
      PORCE = CVS(Mid$(OBX$, 3, 4))
      CADECI% = CVI(Mid$(OBX$, 7, 2))
      FCOR% = CVI(Mid$(OBX$, 9, 2))
      FVIG% = CVI(Mid$(OBX$, 11, 2))
      '
      Screen.MousePointer = 11
      FIN& = NUMAS%
      Call BLANARCH("ACTUCOS")
      I& = 0
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MASTER", U& + 1)
        If CVI(Mid$(X$, 123, 2)) = NCX1% Then
           CIXI% = CVI(Left$(X$, 2))
           CODE$ = CODEXT$(CIXI%)
           If Mid$(CODE$, 16, 1) <> Chr$(96) Then    ' PARA SUPROMIR ITEMS DE BAJA
             MON% = MONECOSTO%(CIXI%)
             PODES1 = 0
             COS1# = COSUNI(CIXI%)
             COS2# = COS1#
             FCOS% = FECOSTO%(CIXI%)
             If FCOS% <= FCOR% Then
                COS2# = ROUND#(COS1# * (1 + PORCE / 100), CADECI%)
                PODES1 = PORCE
             End If
             '
             Y$ = REGBLAN$("ACTUCOS")
             Call REPLA(Y$, MKI$(CIXI%), 1, 2)
             Call REPLA(Y$, MKI$(MON%), 3, 2)
             Call REPLA(Y$, MKD$(COS1#), 5, 8)
             Call REPLA(Y$, MKS$(PODES1), 13, 4)
             Call REPLA(Y$, MKD$(COS2#), 17, 8)
             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
             I& = I& + 1
             Call GRAREG("ACTUCOS", Y$, I&)
           End If
           '
        End If
      Next U&
      Call SETULTREG("ACTUCOS", I&)
      If I& < 1 Then
         Call MENSERR(24, "Proveedor sin Códigos Pre-Asignados")
         GoTo 5
      End If
      '
      Call FILESORT("ACTUCOS", "", 1, 1)
      '
      OPX% = COMALTER%("Opciones de Trabajo:\\Actualizar\Listar")
      If OPX% < 1 Or OPX% > 2 Then GoTo 3
      If OPX% = 2 Then
         Call HEADERS("ACTUCOS", "")
         Call HEADERS("ACTUCOS", "Proveedor: " + RASOCLI$((-1) * NCX1%))
         Call FINAL("*LACTUCOS")
         GoTo 5
      End If
      '
      Call TRALOCAL("ACTUCOS", "")
      Screen.MousePointer = 1
      '
      JUN% = VENTABU%("ACOSPRO/TITUPAN=Costo de Reposición - " + RASOCLI$((-1) * NCX1%))
      If JUN% < 0 Then GoTo 5
      '
      Screen.MousePointer = 11
      For J& = 1 To ULTREG&("!ACTUCOS")
         Z$ = REGLEIDO$("!ACTUCOS", J&)
         CIXI% = CVI(Mid$(Z$, 1, 2))
         COUNUE = CVD(Mid$(Z$, 17, 8))
         
         KU& = CIXI% + 1
         X$ = REGLEIDO$("MASTER", KU&)
         If Abs(COUNUE - CVS(Mid$(X$, 111, 4))) >= 0.00005 Then
           Call REPLA(X$, MKS$(COUNUE), 111, 4)
           Call REPLA(X$, MKI$(FVIG%), 115, 2)
           Call GRAREG("MASTER", X$, KU&)
         End If
      Next J&
      Call CIERRARCH("MASTER")
      Screen.MousePointer = 1
      '
      GoTo 3
    End If
99  Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
3   Call SELECHO("OCOMENCA/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("OCOMENCA"))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
      GoTo 3
    End If
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If DERACCE%("SETUPCOM", "Configuracion de Compras") >= 2 Then
      Call CONECRUN("FLSETUP/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   OPLISCOM.Show 1
   Command6.Enabled = True
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMASTO01"
    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Desanuoc_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOMP.Command10_Click
End Sub

Private Sub DevoConsi_Click()
   OPCICOMP.Command15_Click
End Sub

Private Sub Directa_Click()
   Call Command1_Click
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub Entrecon_Click()
   OPCICOMP.Command2_Click
End Sub

Private Sub Exit_Click()
   Call Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    OPENFORMS = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "ORCOMP02"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If EPAC$ <> "" Then
10     Caption = Caption + "  -  " + EPAC$
       If InStr(EPAC$, "DOSIVAC") > 0 Then
         Text6.Width = 3120
       End If
    End If
    Call GRATITFOR(Caption)
    '
    If ULTREG&("PROVEPLA") > 0 Then
      Command10.Visible = True
    End If
    '
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Nuorig.Enabled = True
      Nuorig1.Enabled = True
    End If
    '
    Show
    DoEvents
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text14.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(PRONUOCO&))
    Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
    '
    Text15.TEXT = TRIM$(CONTROL$("OCOMPRA", "MONECOS"))
    If Text15 = "" Then Text15 = "1"      ' MONEDA DE VALUACIóN
    MONEMI% = 1
    '
    MODOULPRE% = 1
    If CONTROL("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
    '
    DoEvents
    '
    Call SETULTREG("!OCOMDETA", 0)
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label11_DblClick()
'   If TRIM$(UCase$(Label11)) = "PESOS" Then
'       Label11.ForeColor = RGB(255, 0, 0)
'       Label11 = "Dólares"
'       MONEMI% = 2
'       Label11.ToolTipText = "Doble-Click para 'Pesos'"
'       COEFICO = 1 / TICAMONE(2)
'     Else
'       Label11.ForeColor = RGB(0, 0, 255)
'       Label11 = "Pesos"
'       MONEMI% = 1
'       Label11.ToolTipText = "Doble-Click para 'Dólares'"
'       COEFICO = TICAMONE(2)
'   End If
'   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   Label11 = DMONE$(MONEMI%)
   COEFICO = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   Label11.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label11.ForeColor = RGB(255, 0, 0)
   End If
   Label11.ToolTipText = "Doble-Click para Cambiar Moneda"
   '
   If Frame16.Visible = False Then    ' no es de texto libre
        For U& = 1 To ULTREG&("!OCOMDETA")
          YX$ = REGLEIDO$("!OCOMDETA", U&)
          PREDOL# = CVD(Mid$(YX$, 67, 8)) * COEFICO
          Call REPLA(YX$, MKD$(PREDOL#), 67, 8)
          Call GRAREG("!OCOMDETA", YX$, U&)
        Next U&
        Call CIERRARCH("!OCOMDETA")
        Call CARLISCO
        Call CALTOCOM
      Else
        Total(0) = FORMATNUM$(XVALO(Total(0)) * COEFICO, "F11.2") ' ESTO ARRASTRA EL RESTO
   End If
   '
End Sub



Private Sub Label22_DblClick()
   List1.Clear
   Call SETULTREG("!OCOMDETA", 0)
   Unload FORVENTB
   Call CARDETCOM
End Sub

Private Sub List1_Click()
   List1.ToolTipText = DESCRIT$(CODINT%(Left$(List1.TEXT, 15)))
   MUECOPEN% = 1
   For U& = 1 To List2.ListCount
     List2.Selected(U& - 1) = False
     If Left$(List2.List(U& - 1), 15) = Left$(List1.TEXT, 15) Then
       List2.Selected(U& - 1) = True
     End If
   Next U&
   MUECOPEN% = 0
End Sub

Private Sub List1_DblClick()
   Call CARDETCOM
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     Call MUESTRADATCOD(Left$(List1.TEXT, 15))
   End If
   '
End Sub

Private Sub List2_Click()
   If MUECOPEN% < 1 Then
     For U& = 1 To List2.ListCount
       List2.Selected(U& - 1) = False
     Next U&
   End If
End Sub

Private Sub List2_DblClick()
    NPX& = Val(Mid$(List2.TEXT, 79, 6))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
    End If
End Sub

Private Sub Ltota_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
     If Total(KKI%).Enabled = True Then
          Total(KKI%).BackColor = &H80000005
       Else
          Total(KKI%).BackColor = &HE0E0E0
     End If
   Next KKI%
End Sub

Private Sub MarItCum_Click()
   Call OPCICOMP.Command4_Click
End Sub

Private Sub Master_Click()
   Command8_Click
End Sub

Private Sub Modi_Ocom_Click()
   Call SETULTREG("!OCOMDETA", 0)
   Call BLANFORMU
   Call OPCICOMP.Command13_Click
End Sub

Private Sub Nuorig_Click()
   Call OPCICOMP.Command19_Click
End Sub

Private Sub Nuorig1_Click()
   Call OPCICOMP.Command19_Click
End Sub

Private Sub Por_MRP_Click()
   Call Command10_Click
End Sub

Private Sub PORDESCU_Change()
   Call CALTOCOM
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Print1_Click()
   OPLISCOM.Option2.Value = True
   Call OPLISCOM.Command3_Click
End Sub

Private Sub Print2_Click()
   OPLISCOM.Option2.Value = True
   Call OPLISCOM.Command4_Click
End Sub

Private Sub Print3_Click()
   OPLISCOM.Option2.Value = True
   Call OPLISCOM.Command5_Click
End Sub

Private Sub Print4_Click()
   OPLISCOM.Option2.Value = True
   Call OPLISCOM.Command9_Click
End Sub

Private Sub Provedor_Click()
   Command11_Click
End Sub

Private Sub Screen1_Click()
   OPLISCOM.Option1.Value = True
   Call OPLISCOM.Command3_Click
End Sub

Private Sub Screen2_Click()
   OPLISCOM.Option1.Value = True
   Call OPLISCOM.Command4_Click
End Sub

Private Sub Screen3_Click()
   OPLISCOM.Option1.Value = True
   Call OPLISCOM.Command5_Click
End Sub

Private Sub Screen4_Click()
   OPLISCOM.Option1.Value = True
   Call OPLISCOM.Command9_Click
End Sub

Private Sub Setup_Click()
  OPCICOMP.Command18_Click
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("PROVEPLA")
  NCLIE% = Val(TRIM$(VALACT1$("PROVEPLA")))
  If NCLIE% > 0 Then
    If REGICLI&((-1) * NCLIE%) > 0 Then
      If Frame16.Visible = True Then
        Text1.TEXT = TRIM$(Str$(NCLIE%))
        Call CARDACLI
        Text28 = ""
        Text18 = ""
        Exit Sub
      End If
      '
      Call BLANFORMU
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      Call CARDACLI
      Call CARCOPEN
      Call CARCOANT
      If ULTREG&("!OCOMDETA") > 0 Then
           Call CARLISCO
           Call CALTOCOM
        Else
           Call CARDETCOM
      End If
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   '
   If TRIM$(Text1.TEXT) <> "" Then
      Command1.SetFocus
      Exit Sub
   End If
   '
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   '
   Label11.ForeColor = RGB(0, 0, 255)
   Label11 = "Pesos"
   MONEMI% = 1
   Label11.ToolTipText = "Doble-Click para 'Dólares'"
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      Call CARDACLI
      Call CARCOPEN
      Call CARCOANT
      If ULTREG&("!OCOMDETA") > 0 Then
           Call CARLISCO
           Call CALTOCOM
         Else
           Call CARDETCOM
      End If
    End If
  End If
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If ULTREG&("PEDENCA") > 2 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
   '
   NPP& = PRONUOCO&
   If Val(NUMEPRE$) < NPP& Then
     Call MENSERR(24, "Número no Válido")
     Text11.TEXT = TRIM$(Str$(NPP&))
     Exit Sub
   End If
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Presupuesto Ajustado")
   End If
   '
End Sub

Private Sub Text12_Change()
   Call CALTOCOM
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text15_Change()
   Text16 = ECOARCH$("TAMONED", Chr$(Val(Text15)))
End Sub

Private Sub Text17_Change()
   Call CALTOCOM
End Sub

Private Sub Text18_Change()
   CIXI% = CODINT%(Text18)
   Label8 = DESCRIT0$(CIXI%)
   Call CARESPECI
End Sub

Sub CARESPECI()
      '
    CODD$ = Text18.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
      TTYZ$ = ""
      GoTo 30
    End If
    '
    Text28.TEXT = "": TTYZ$ = ""
    If CI% > 0 Then
       If CI% <= NUMAS% Then
          Screen.MousePointer = 11
          RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
          URN& = ULTREG&("ESPECIFI")
          '
          For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
                If RENOTA& <= URN& Then
                   TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   If TTYZ$ <> "" Then
                      TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
                   End If
                   TTYZ$ = TTYZ$ + RTrim$(TTYY$)
                End If
             End If
          Next U&
          Screen.MousePointer = 1
       End If
    End If
    LOCADE& = Len(TTYZ$)
    For U& = LOCADE& To 1 Step -1
      If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
    Next U&
    U& = 0
    '
30  Text28.TEXT = Left$(TTYZ$, U&)
    Text28.Refresh
    '
End Sub
'

Private Sub Text18_DblClick()
   Call Command18_Click
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If Val(TEXT21(Index).TEXT) < 0 Or Val(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", Chr$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("TAEMBALA", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "CONPAG": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "TAEMBALA": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If ECOARCH$("PROVED1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Call SETULTREG("!OCOMDETA", 0)
   Call CIERRARCH("!OCOMDETA")
   Call BLANARCH("!OCOMDETA")
   Call SETULTREG("!1OCOMDET", 0)
   Call CIERRARCH("!1OCOMDET")
   Call BLANARCH("!1OCOMDET")
   On Error Resume Next
   Unload FORVENTB
   '
   Label5 = ""
   Text18 = ""
   Text28 = ""
   Frame16.Visible = False
   Check1.Value = 0
   Frame2.Top = 3960
   Frame9.Visible = True
   '
   On Error GoTo 0
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text12.TEXT = FORMATNUM$(ALIVA, "F5.2")
   Text13.TEXT = ""
   Text6.TEXT = ""
   List1.Clear: List1.ToolTipText = ""
   List2.Clear
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   'TEXT21(2).TEXT = "": Eco21(2) = ""
   '
   PORDESCU.TEXT = ""
   Text17 = ""
   For KKI% = 0 To 4
      Total(KKI%).Enabled = True
      Total(KKI%).TEXT = ""
      If Total(KKI%).Enabled = True Then
          Total(KKI%).BackColor = &H80000005
        Else
          Total(KKI%).BackColor = &HE0E0E0
      End If
   Next KKI%
   '
   'MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUOCO&))
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = (-1) * Val(Text1.TEXT)
    If NC% < 0 Then
      If ECOARCH$("ACREDOR", MKI$((-1) * NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text7.TEXT = TRIM$(POSIVAC$(NC%))
        Text8.TEXT = TRIM$(CUITCLI$(NC%))
        '
        TEXT21(0).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(1).TEXT = ""
        '
        Text15 = TRIM$(Str$(MONEVACO%(NC%)))
        Label11.ForeColor = RGB(0, 0, 255)
        Label11 = "Pesos"
        MONEMI% = 1
        Label11.ToolTipText = "Doble-Click para 'Dólares'"
        If MONEMICO%(NC%) = 2 Then
          Label11.ForeColor = RGB(255, 0, 0)
          Label11 = "Dólares"
          MONEMI% = 2
          Label11.ToolTipText = "Doble-Click para 'Pesos'"
        End If
        '
        ATEN$ = "": PODESC$ = ""
        RFF$ = RECFILT$("OCOMENCA", 4, MKI$(Abs(NC%)))
        For KKK% = Len(RFF$) - 3 To 1 Step -4
          REOCO& = CVS(Mid$(RFF$, KKK%, 4))
          XX$ = REGLEIDO$("OCOMENCA", REOCO&)
          ATEN$ = TRIM$(Mid$(XX$, 163, 30))
          PODESC$ = TRIM$(Mid$(XX$, 97, 16))
          If ATEN$ <> "" Then Exit For
        Next KKK%
        Text6 = ATEN$
        PORDESCU = PODESC$
        '
        If EXISTE%(LUDAT$ + "CALINTEG.RFS") > 0 Then
          XX11$ = FILTERGETRECORD$("CALINTEG", 1, MKI$((-1) * NC%))
          PUNTAPRO = CVS(Mid$(XX11$, 35, 4))
          Label5 = FORMATNUM$(PUNTAPRO, "F4.1")
        End If
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
   '
Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Private Sub Text6_GotFocus()
   On Error Resume Next
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
   '
   'Call CALTOCOM
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOCOM
'      DESC$ = PORDESCU.TEXT
'      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = Val(Total(2).TEXT)
'      TTT2 = Val(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub

Private Sub Total_DblClick(Index As Integer)
   For KKI% = 0 To 4
     Total(KKI%).Enabled = Not (Total(KKI%).Enabled)
     If Total(KKI%).Enabled = True Then
          Total(KKI%).BackColor = &H80000005
       Else
          Total(KKI%).BackColor = &HE0E0E0
     End If
   Next KKI%
End Sub
'
Private Sub Total_GotFocus(Index As Integer)
   If Frame16.Visible = True Then
     If VALICLIE% > 0 Then
       If CODINT%(Text18) > 0 Then
         If Index < 1 Then
           Exit Sub
         End If
       End If
     End If
   End If
   '
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRACOMPRA(OKX%)
   '
   If Frame16.Visible = True Then
     ZZ$ = REGBLAN$("OCOMDETA")
       CIXI% = CODINT%(Text18)
       If CIXI% < 1 Then
         MERRO$ = "Falta Codigo de Articulo"
         OKX% = 0
         GoTo 98
       End If
     Call REPLA(ZZ$, MKI$(CIXI%), 9, 2)
       UNISTO$ = TRIM$(UNIMED$(CIXI%))
       If UNISTO$ = "" Then UNISTO$ = "U."
     Call REPLA(ZZ$, UNISTO$, 11, 4)
     Call REPLA(ZZ$, MKS$(1), 51, 4)
     Call REPLA(ZZ$, MKS$(1), 55, 4)
     Call REPLA(ZZ$, UNISTO$, 59, 4)
       If Val(Total(0)) < 0.005 Then
         MERRO$ = "Falta Precio de Compra"
         OKX% = 0
         GoTo 98
       End If
     Call REPLA(ZZ$, MKD$(Val(Total(0))), 67, 8)
     Call REPLA(ZZ$, MKS$(0), 75, 4)
       FENTRESO% = FECHANUM(Text14)
     Call REPLA(ZZ$, MKI$(FENTRESO%), 95, 2)
     Call GRAREG("!OCOMDETA", ZZ$, 1)
     Call SETULTREG("!OCOMDETA", 1)
     Call CIERRARCH("!OCOMDETA")
     Frame16.Visible = False
     DoEvents
     Call CALTOCOM
     Frame16.Visible = True
     DoEvents
   End If
   '
   OKX% = 0
   If ULTREG&("!OCOMDETA") < 1 Then Exit Sub
   '
   COPRECIO = 1
   If Total(0).Enabled = False Then COPRECIO = 0
   '
   Call CIERRARCH("OCOMENCA")
   Call CIERRARCH("OCOMDETA")
   '
   Call BLOQARCH("OCOMENCA")
   Call BLOQARCH("OCOMDETA")
   Call BLOQARCH("MACONSIG")
   '
   NPP& = PRONUOCO&
   If NPP& > Val(Text11.TEXT) Then
      Call CIERRARCH("OCOMENCA")
      Call CIERRARCH("OCOMDETA")
      Call CIERRARCH("MACONSIG")
      Call COMUNI("FLEXOFT ha Re-Numerado la O/Compra\como O/C Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   MONECO% = Val(Text15)
   If MONECO% < 1 Then MONECO% = Val(TRIM$(CONTROL$("OCOMPRA", "MONECOS")))
   If MONECO% < 1 Then MONECO% = 1
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Proveedor no Válido"
      GoTo 98
   End If
   NPRO% = Val(Text1.TEXT)
   '
   If Val(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If Val(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Embalaje no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
123 If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
    End If
    '
    ENCAO$ = REGBLAN$("OCOMENCA")
    '
      NRO& = Val(Text11)
    Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
      FEX = FEMI%
      REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
    Call REPLA(ENCAO$, REFE$, 5, 44)
    Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
    Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
    Call REPLA(ENCAO$, Chr$(0), 53, 1)
    Call REPLA(ENCAO$, Chr$(0), 54, 1)
      MONECOI% = Val(Text15)
    Call REPLA(ENCAO$, Chr$(MONECOI%), 83, 1)
    Call REPLA(ENCAO$, Chr$(MONEMI%), 84, 1)
      TICAMBX = 1
      If MONEMI% <= 1 Then
        If MONECOI% > 1 Then
          TICAMBX = TICAMONE(MONECOI%)
        End If
      End If
    Call REPLA(ENCAO$, MKS$(TICAMBX), 85, 4)
    '
      HO% = HOY(HOS$)
      HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
    Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
    Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
      '
      CPAG% = Val(TEXT21(0))
    Call REPLA(ENCAO$, MKI$(CPAG%), 153, 2)
    '
      ATEN$ = TRIM$(Text6)
    Call REPLA(ENCAO$, ATEN$, 163, 30)
    '
125 PORDESCUENTO$ = TRIM$(PORDESCU)
    DESC$ = PORDESCUENTO$
    TOTANE# = Val(Total(0)) * COPRECIO
    IDESCUEN# = Val(Total(1)) * COPRECIO
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = Val(Total(2)) * COPRECIO
    TOTPER# = Val(Total(4)) * COPRECIO
    TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
    '
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, TRIM$(PORDESCU), 97, 16)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTPER#), 129, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    LUENTREGA$ = Text10
    Call REPLA(ENCAO$, LUENTREGA$, 449, 64)
    '
    Call FRATEXTO(Text13, 52)
    For KKU& = 1 To CARETEX%
      If KKU& <= 4 Then
        Call REPLA(ENCAO$, RETEX$(KKU&), 129 + 64 * KKU&, 64)
      End If
    Next KKU&
    Screen.MousePointer = 11
    '
    RENCA& = 1 + ULTREG&("OCOMENCA")
    PRIDETA& = 1 + ULTREG&("OCOMDETA")
    ULTDETA& = PRIDETA& + ULTREG&("!OCOMDETA") - 1
    '
    ' grabar encabezado
    Call REPLA(ENCAO$, MKS$(PRIDETA&), 155, 4)
    Call REPLA(ENCAO$, MKS$(ULTDETA&), 159, 4)
    Call GRAREG("OCOMENCA", ENCAO$, RENCA&)
    Call CIERRARCH("OCOMENCA")
    '
    EMICONSI% = 0
    IDENOC$ = MKS$(NRO&) + MKI$(FEMI%) + MKI$(NPRO%)
    For J& = 1 To ULTREG&("!OCOMDETA")
       ZZ$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(ZZ$, IDENOC$, 1, 8)
       Call REPLA(ZZ$, MKS$(RENCA&), 125, 4)
       Call REPLA(ZZ$, Chr$(MONECOI%) + Chr$(MONEMI%), 121, 2)
       If COPRECIO < 0.99 Then
         Call REPLA(ZZ$, String$(28, 0), 67, 28)  ' BORRA PRECIO UNITARIO E IMPORTE
         Call GRAREG("!OCOMDETA", ZZ$, J&)
       End If
         '
         COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
         PREUNPE# = CVD(Mid$(ZZ$, 67, 8)) * TICAMONE(MONEMI%) / COEFUNI
         PREUNCO# = PREUNPE# / TICAMONE(MONECOI%)
         MONECOS$ = "$ ": If MONECOI% > 1 Then MONECOS$ = Left$(Text16, 2)
         VAUNI$ = MONECOS$ + Mid$(FORMATNUM$(PREUNCO#, "F12.4"), 3)
         '
       Call REPLA(ZZ$, VAUNI$, 43, 12)
       Call REGAPP("OCOMDETA", ZZ$)
       '
       NOCOM$ = TRIM$(Str$(NRO&))
       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
       ITNU$ = TRIM$(Str$(J&))
       While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
       '
       CI0% = CVI(Mid$(ZZ$, 9, 2))
       CANTU = CVS(Mid$(ZZ$, 55, 4))
       NPX$ = MKI$(NPRO%)
       Call LEEEXPLO(CODEXT$(CI0%), CANTU)
       For I% = 1 To CAIT%
          YY$ = REGBLAN$("MACONSIG")
          Call REPLA(YY$, MKI$(FECHEMI%), 1, 2)
          Call REPLA(YY$, NPX$, 3, 2)
          Call REPLA(YY$, NOCOM$, 21, 12)
          '
          CII% = CIJ%(I%)
          CANTI# = CDbl(CANTU * USOI(I%))
          '
          If TRIM$(CODEXT$(CII%)) <> "" Then
            If CANTI# > 0.05 Then
              Call REPLA(YY$, MKI$(CII%), 33, 2)
              Call REPLA(YY$, UNIMED$(CII%), 51, 4)
              Call REPLA(YY$, MKD$(CANTI#), 55, 8)
              Call REGAPP("MACONSIG", YY$)
              EMICONSI% = 1
            End If
          End If
       Next I%
       '
    Next J&
    Call CIERRARCH("OCOMDETA")
    Call CIERRARCH("MACONSIG")
    '
    If Frame16.Visible = True Then
       TTXX$ = REGBLAN$("TXOCOLIB")
       Call REPLA(TTXX$, MKS$(NRO&), 1, 4)
          DEOC$ = TRIM$(Text28)
          If DEOC$ = "" Then DEOC$ = TRIM$(Label8)
       Call REPLA(TTXX$, AJUSTI$(DEOC$, 64), 5, 64)
       Call REGAPP("TXOCOLIB", TTXX$)
       Call CIERRARCH("TXOCOLIB")
    End If
    Call CIERRARCH("TXOCOLIB")
    Call BAJALDISCO
    '
    If COPRECIO < 0.99 Then
      PORDESCU.TEXT = ""
      Text17 = ""
      For KK% = 0 To 4
        Total(KK%).TEXT = ""
      Next KK%
    End If
    '
    ' actualizar costo standard y kilos por pieza
    ACOSREP$ = UCase$(CONTROL$("OCOMPRA", "ACOSREP"))
    'If ACOSREP$ = "SIEMPRE" Then Option7.Value = True
    'If ACOSREP$ = "NUNCA" Then Option9.Value = True
    '
    HOII% = HOY(HOS$)
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        COSPESOSD# = CVD(Mid$(ZZ$, 79, 8)) * TICAMONE(MONEMI%) / COEFUNI
        '
        REMAL& = CI0% + 1
        REMAS$ = REGLEIDO$("MASTER", REMAL&)
        PRHBL% = CVI(Mid$(REMAS$, 123, 2))
        If PRHBL% < 1 Or PRHBL% = NPRO% Or ACOSREP$ = "SIEMPRE" Then
            COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
            If COSMONCO >= 0.000005 Then
              If ACOSREP$ <> "NUNCA" Then
                Call REPLA(REMAS$, MKS$(COSMONCO), 111, 4)
                Call REPLA(REMAS$, MKI$(FEMI%), 115, 2)
                If PRHBL% < 1 Then Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
                Call REPLA(REMAS$, Chr$(MONECO%), 95, 1)
                Call GRAREG("MASTER", REMAS$, REMAL&)
              End If
            End If
            '
          Else
            '
            TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(CI0%)) + "\el Proveedor Habitual Registrado es\"
            TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PRHBL%)) + "\  \"
            TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
            TTXX$ = TTXX$ + RASOCLI$((-1) * NPRO%)
            If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 2 Then
              '
              Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
              COSMONCO = COSPESOSD# / TICAMONE(MONECO%)
              If COSMONCO >= 0.000005 Then
                Call REPLA(REMAS$, MKS$(COSMONCO), 111, 4)
                Call REPLA(REMAS$, MKI$(FEMI%), 115, 2)
                Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
                Call REPLA(REMAS$, Chr$(MONECO%), 95, 1)
              End If
              Call GRAREG("MASTER", REMAS$, REMAL&)
              '
              ZZX$ = REGBLAN$("RECAMPH")
              Call REPLA(ZZX$, MKI$(CI0%), 1, 2)
              HOII% = HOY(HOS$)
              Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
              Call REPLA(ZZX$, MKI$(PRHBL%), 5, 2)
              Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
              Call REPLA(ZZX$, USERTER$, 9, 24)
              Call REPLA(ZZX$, "Emis.O/C " + TRIM$(Str$(NRO&)), 33, 32)
              Call REGAPP("RECAMPH", ZZX$)
              Call CIERRARCH("RECAMPH")
              GoTo 169
              '
            End If
            '
            CUNIREG# = COSUNI(CI0%) * TICAMONE(MONECOSTO(CI0%))
            If CUNIREG# >= 0.00005 Then
              If COSPESOSD# > CUNIREG# + 0.00005 Then
                Call COMUNI("ATENCION: Precio de Compra\es Superior al Costo Standard\en Código '" + TRIM$(CODEXT$(CI0%)) + "' !!!\  \Informe al Supervisor de Compras.")
              End If
            End If
169     End If
        '
        If PESUNI(CI0%) < 0.000005 Then
          UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
          If Left$(UNIST$, 1) = "U" Then
            UNICO$ = UCase$(TRIM$(Mid$(ZZ$, 59, 4)))
            If UNICO$ = "KG" Then
              On Error Resume Next
              PSXX = 1 / COEFUNI
              On Error GoTo 0
              REMAL& = CI0% + 1
              REMAS$ = REGLEIDO$("MASTER", REMAL&)
              Call REPLA(REMAS$, MKS$(PSXX), 99, 4)
              Call GRAREG("MASTER", REMAS$, REMAL&)
            End If
          End If
        End If
        '
      End If
    Next J&
    '
    '
    'depurar PLACOMP
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        COEFUNI = CVS(Mid$(ZZ$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        CANOCOM = CVS(Mid$(ZZ$, 51, 4))
        FENTRESO% = CVI(Mid$(ZZ$, 95, 2))
        If FENTRESO% < FEMI% Then FENTRESO% = FEMI%
        '
        RFF$ = RECFILT$("PLACOMP", 1, MKI$(CI0%))
        For UJ& = 1 To Len(RFF$) Step 4
          RPLA& = CVS(Mid$(RFF$, UJ&, 4))
          XX$ = REGLEIDO$("PLACOMP", RPLA&)
          FENTREPLA% = CVI(Mid$(XX$, 3, 2))
            If FENTREPLA% < FEMI% Then FENTREPLA% = FEMI%
          CANPLA = CVS(Mid$(XX$, 21, 4))
          '
          If FENTREPLA% >= FENTRESO% Then
            If CANPLA <= CANOCOM Then
                CANOCOM = CANOCOM - CANPLA
                CANPLA = 0
              Else
                CANPLA = CANPLA - CANOCOM
                CANOCOM = 0
            End If
            Call REPLA(XX$, MKS$(CANPLA), 21, 4)
            ' depurar registro de plan de compras
            If CANPLA < 0.005 Then XX$ = String$(128, 0)
            Call GRAREG("PLACOMP", XX$, RPLA&)
          End If
        Next UJ&
      End If
    Next J&
    Call CIERRARCH("PLACOMP")
    '
    JJ& = 0
    For U& = 193 To 512 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      JJ& = JJ& + 1
      Call GRAREG("!OBSERVOF", TTXX$, JJ&)
    Next U&
    Call SETULTREG("!OBSERVOF", JJ&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = CVI(Mid$(ENCAO$, 153, 2))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
    Screen.MousePointer = 1
    OKX% = 1
    '
    Exit Sub
    '
98  Call CIERRARCH("OCOMENCA")
    Call CIERRARCH("OCOMDETA")
    Call CIERRARCH("MACONSIG")
    Screen.MousePointer = 1
    Call MENSERR(24, MERRO$)
    '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = Val(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETCOM()
   '
   Call CIERRARCH("!OCOMDETA")
   List2.ListIndex = (-1)
   HOII% = HOY(HOS$)
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Orden de Compra en " + TRIM$(Label11)
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   PRF$ = "CARDETCO"
   XXX% = VENTABU%(PRF$ + ATRI$)
   List1.Clear
   Call CALTOCOM
   Timer1.Enabled = False
   '
   If XXX% < 0 Or ULTREG&("!OCOMDETA") < 1 Then
      Call CIERRARCH("!OCOMDETA")
      Call BLANARCH("!OCOMDETA")
      Call BLANFORMU
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   ATK% = 0
   NPRO% = Val(Text1)
   For U& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", U&)
      Z0$ = ZZ$
      CI0% = CVI(Mid$(ZZ$, 9, 2))
      If TRIM$(Mid$(ZZ$, 11, 4)) = "" Then Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
      If TRIM$(Mid$(ZZ$, 59, 4)) = "" Then Call REPLA(ZZ$, UNICOM$(CI0%), 59, 4)
      If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(ZZ$, 59, 4)) Then Call REPLA(ZZ$, MKS$(1), 63, 4)
      '
      UC$ = ""
      COEFI = CVS(Mid$(ZZ$, 63, 4))
      CANCOMP = CVS(Mid$(ZZ$, 55, 4))
      CANSTOK = CVS(Mid$(ZZ$, 51, 4))
      If COEFI < 0.0001 Then
        If CANCOMP > 0.05 Then
          If CANSTOK > 0.05 Then
            COEFI = CANSTOK / CANCOMP
          End If
        End If
      End If
      '
      If COEFI < 0.0001 Then
        UC$ = ULTICOM$(NPRO%, CI0%)
        If TRIM$(Mid$(ZZ$, 11, 4)) = TRIM$(Mid$(UC$, 11, 4)) Then
          If TRIM$(Mid$(ZZ$, 59, 4)) = TRIM$(Mid$(UC$, 59, 4)) Then
            COEFI = CVS(Mid$(UC$, 63, 4))
          End If
        End If
      End If
      If COEFI < 0.0001 Then COEFI = 1
      '
      If CANCOMP < 0.05 Then
          CANCOMP = CANSTOK / COEFI
        Else
          CANSTOK = CANCOMP * COEFI
      End If
      CANCOMP = CANSTOK / COEFI
      Call REPLA(ZZ$, MKS$(CANSTOK), 51, 4)
      Call REPLA(ZZ$, MKS$(CANCOMP), 55, 4)
      Call REPLA(ZZ$, MKS$(COEFI), 63, 4)
      '
      PREUNID# = CVD(Mid$(ZZ$, 67, 8))
      If PREUNID# < 0.00005 Then
        If UC$ = "" Then UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PREUNID# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        PORDESCUEI = CVS(Mid$(UC$, 75, 4))
        If PREUNID# < 0.00005 Or MODOULPRE% = 2 Then
          PREUNID# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFI
          PORDESCUEI = 0
        End If
        PREUNID# = PREUNID# / TICAMONE(MONEMI%)
        Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
        Call REPLA(ZZ$, MKS$(PORDESCUEI), 75, 4)
      End If
      '
      If ZZ$ <> Z0$ Then ATK% = 1
      Call GRAREG("!OCOMDETA", ZZ$, U&)
   Next U&
   '
   Screen.MousePointer = 1
   If ATK% = 1 Then GoTo 10
   '
   Call CARLISCO
   Call CALTOCOM
   Screen.MousePointer = 1
   '
End Sub

Sub CALTOCOM()
   '
   NC1% = (-1) * Val(Text1)
   TOTANE# = 0
   If Frame16.Visible = True Then
     TOTANE# = Val(Total(0))
     GoTo 20
   End If
   '
   For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PREUNID# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       '
       If CANTU > 0.05 Then
         COEFI = CANSTXX / CANTU
         Call REPLA(X$, MKS$(COEFI), 63, 4)
         Call GRAREG("!OCOMDETA", X$, J&)
       End If
       '
   Next J&
   '
   Call CIERRARCH("!OCOMDETA")
   '
' presentar registro de totales y condicion de pago
   '
20 If MONEMI% > 1 Then
     MONECOI% = MONEMI%
     TICAMBX = TICAEMI
   End If
   '
   PORDESCUENTO$ = TRIM$(PORDESCU)
   DESC$ = PORDESCUENTO$
   IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
   TOTANEDES# = TOTANE# - IDESCUEN#
   TOTIVA# = 0
   ALIVX = Val(Text12)
   If ALIVX < 0 Or ALIVX > 99 Then
      Text12 = TRIM$(Str$(ALIVA))
      ALIVX = ALIVA
   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
           TOTIVA# = (Int(ALIVX * TOTANEDES# + 0.5)) / 100
         End If
       End If
     End If
   End If
   '
   TOTPER# = 0
   PERCX = Val(Text17)
   If PERCX < 0 Or PERCX > 99 Then
      Text17 = ""
      PERCX = 0
   End If
   If Abs(NC1%) > 0 Then
     PVCL% = PIVACLI%(NC1%)
     If PVCL% <> 2 Then
       If PVCL% <> 4 Then
         If PVCL% <> 5 Then
           TOTPER# = (Int(PERCX * TOTANEDES# + 0.5)) / 100
         End If
       End If
     End If
   End If
   '
   
   TOTATO# = TOTANEDES# + TOTIVA# + TOTPER
   '
   If Frame16.Visible = False Then
     Total(0) = FORMATNUM$(TOTANE#, "F11.2")
   End If
   Total(1) = FORMATNUM$(IDESCUEN#, "F11.2")
   Total(2) = FORMATNUM$(TOTIVA#, "F11.2")
   Total(4) = FORMATNUM$(TOTPER#, "F11.2")
   Total(3) = FORMATNUM$(TOTATO#, "F11.2")
   '
End Sub

Sub CARCOPEN()
   '
   Screen.MousePointer = 11
   NPRO% = Val(Text1)
   HO% = HOY(HOS$)
   '
   JJ& = 0
   REBLA$ = REGBLAN$("OCOMDETA")
   Call SETULTREG("!OCOMDETA", 0)
   ARCHILIB$ = "!1OCOMDET"
   Call SETULTREG(ARCHILIB$, 0)
   '
   PCI$ = "": UFECHA = HOY(HOS$)
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   If RFF$ = "" Then Exit Sub
   '
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% > UFECHA Then UFECHA = FENTREIT%
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 9
      Next UKI%
      PCI$ = PCI$ + MCI0$
9  Next KKI%
   '
   FORSELMAT.LIMATSEL.Clear
   For UKI% = 1 To Len(PCI$) Step 2
     CI0% = CVI(Mid$(PCI$, UKI%, 2))
     TTXX$ = AJUSTI$(CODEXT$(CI0%), 16) + DESCRIT0$(CI0%)
     FORSELMAT.LIMATSEL.AddItem TTXX$
   Next UKI%
   FORSELMAT.Caption = "Materiales a Comprar - " + Text2
   FORSELMAT.Label2 = FECHATEX$(UFECHA)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     FORSELMAT.LIMATSEL.Selected(U& - 1) = True
   Next U&
   '
   Screen.MousePointer = 1
   On Error Resume Next
   FORSELMAT.BOTREC.SetFocus
   FORSELMAT.LIMATSEL.TopIndex = 0
   On Error GoTo 0
   FORSELMAT.Show 1
   '
   DoEvents
   If TRIM$(FORSELMAT.Label2 = "") Then Exit Sub
   '
   Screen.MousePointer = 11
   PCI$ = "":
   UFECHAI% = FECHANUM(FORSELMAT.Label2)
   For U& = 1 To FORSELMAT.LIMATSEL.ListCount
     If FORSELMAT.LIMATSEL.Selected(U& - 1) = True Then
       PCI$ = PCI$ + MKI$(CODINT%(TRIM$(Left$(FORSELMAT.LIMATSEL.List(U& - 1), 15))))
     End If
   Next U&
   '
   RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPX& = CVS(Mid$(RFF$, KKI%, 4))
      XX$ = REGLEIDO$("PLACOMP", REPX&)
      CI0% = CVI(Left$(XX$, 2))
      MCI0$ = MKI$(CI0%)
      For UKI% = 1 To Len(PCI$) Step 2
        If Mid$(PCI$, UKI%, 2) = MCI0$ Then GoTo 24
      Next UKI%
      GoTo 29
      
24    FENTREIT% = CVI(Mid$(XX$, 3, 2))
      If FENTREIT% < HO% Then FENTREIT% = HO%
      If FENTREIT% <= UFECHAI% Then
        '
        CANTU = CVS(Mid$(XX$, 21, 4))
        '
        For KKJ& = 1 To JJ&
          ZZ$ = REGLEIDO$("!OCOMDETA", KKJ&)
          If CVI(Mid$(ZZ$, 9, 2)) = CI0% Then
            If CVI(Mid$(ZZ$, 95, 2)) = FENTREIT% Then
              CANTU = CANTU + CVS(Mid$(ZZ$, 51, 4))
              Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
              Call GRAREG("!OCOMDETA", ZZ$, KKJ&)
              GoTo 29
            End If
          End If
        Next KKJ&
        '
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
        Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
        Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
        Call REPLA(ZZ$, UNIMED$(CI0%), 59, 4)
        Call REPLA(ZZ$, MKS$(1), 63, 4)
        Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
        '
        UC$ = ULTICOM$(NPRO%, CI0%)
        MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
        PREUNID# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
        COEFIX = CVS(Mid$(UC$, 63, 4)): If COEFIX < 0.0005 Then COEFIX = 1
        If PREUNID# < 0.00005 Or MODOULPRE% = 2 Then
          PREUNID# = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * COEFIX
        End If
        PREUNID# = PREUNID# / TICAMONE(MONEMI%)
        '
        UNICOMPRA$ = TRIM$(Mid$(UC$, 59, 4))
        If UNICOMPRA$ <> "" Then Call REPLA(ZZ$, UNICOMPRA$, 59, 4)
        COEFUNI = CVS(Mid$(UC$, 63, 4)): If COEFUNI < 0.0001 Then COEFUNI = 1
        If UNICOMPRA$ = TRIM$(UNIMED$(CI0%)) Then COEFUNI = 1
        Call REPLA(ZZ$, MKS$(COEFUNI), 63, 4)
        
        Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
        'End If
        JJ& = JJ& + 1
        Call GRAREG("!OCOMDETA", ZZ$, JJ&)
          DELIBX$ = DESCRIT0$(CI0%)
        Call GRAREG(ARCHILIB$, DELIBX$, JJ&)
      End If
29 Next KKI%
   '
   Call SETULTREG("!OCOMDETA", JJ&)
   Screen.MousePointer = 1
   '
End Sub

Sub CARCOANT()
   '
   Screen.MousePointer = 11
   List2.Clear
   NPRO% = Val(Text1)
   HOU = HOY(HOS$)
   '
   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
   For UK& = 1 To Len(RFF$) Step 4
     U& = CVS(Mid$(RFF$, UK&, 4))
     X$ = REGLEIDO$("OCOMDETA", U&)
     STAT% = Asc(Mid$(X$, 124, 1))
     If STAT% < 8 Then      ' no esta anulada ni marcada cumplida
       NUORSE& = CVS(Mid$(X$, 117, 4))
       If NUORSE& < 1 Then
         COEFCAN = CVS(Mid$(X$, 63, 4))
         If COEFCAN < 0.005 Then COEFCAN = 1
         CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
         CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
         CAPEN = CACOM - CAREC
         If CAPEN / COEFCAN >= 0.05 Then
           CIXI% = CVI(Mid$(X$, 9, 2))
           NROPED& = CVS(Left$(X$, 4))
           UNICO$ = TRIM$(Mid$(X$, 59, 4))
           FEX = CVI(Mid$(X$, 95, 2))
           ATRAU = 0
           If FEX < HOU Then
             ATRAU = (DIENTRE(FEX, HOU)) / 7
             If ATRAU > 99.9 Then ATRAU = 99.9
           End If
           ATRAX$ = FORMATNUM$(ATRAU, "F4.1")
           '
           LI2TEX$ = Space$(80)
           Call REPLA(LI2TEX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(LI2TEX$, DESCRIT0$(CIXI%), 17, 33)
           Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN / COEFCAN), 4, 8, 1, 2), 51, 8)
           Call REPLA(LI2TEX$, UNICO$, 60, 4)
           Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 8)
           Call REPLA(LI2TEX$, ATRAX$, 74, 4)
           Call REPLA(LI2TEX$, FORMATNUM$(NROPED&, "I6"), 79, 6)
           List2.AddItem LI2TEX$
           '
         End If
       End If
     End If
   Next UK&
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISCO()
   '
   Screen.MousePointer = 11   ' cargando Lista de Compra
   HOII% = HOY(HOS$)
   FENTREGEN% = FECHANUM(Text14)
   List1.Clear: List1.ToolTipText = ""
   FIN& = ULTREG&("!OCOMDETA")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   ARCHILIB$ = "!1OCOMDET"
   '
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!OCOMDETA", U&)
     CI% = CVI(Mid$(X$, 9, 2))
     FENTRESOL% = CVI(Mid$(X$, 95, 2))
     If FENTRESOL% < FENTREGEN% Then FENTRESOL% = FENTREGEN%
     If FENTRESOL% < HOII% Then FENTRESOL% = HOII%
     If FENTRESOL% <> CVI(Mid$(X$, 95, 2)) Then
        Call REPLA(X$, MKI$(FENTRESOL%), 95, 2)
        Call GRAREG("!OCOMDETA", X$, U&)
     End If
     '
     UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
     UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
     CANSTOK = CVS(Mid$(X$, 51, 4))
     CANCOMP = CVS(Mid$(X$, 55, 4))
     'If CANSTOK < 0.05 Then
     '   Call MENSERR(24, "Falta Cantidad en " + UNIST$ + "\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
     'End If
     '
     COEFI = CVS(Mid$(X$, 63, 4))
     If UNIST$ = UNICO$ Then COEFI = 1
     If COEFI < 0.0001 Then COEFI = 1
     '
     If CANCOMP < 0.05 Then
         CANCOMP = CANSTOK / COEFI
       Else
         CANSTOK = CANCOMP * COEFI
     End If
     CANCOMP = CANSTOK / COEFI
     Call REPLA(X$, MKS$(CANSTOK), 51, 4)
     Call REPLA(X$, MKS$(CANCOMP), 55, 4)
     Call REPLA(X$, MKS$(COEFI), 63, 4)
     Call GRAREG("!OCOMDETA", X$, U&)
     '
     UNISTO$ = Mid$(X$, 11, 4)
     UNICOX$ = Mid$(X$, 59, 4)
     COEFUNI = CVS(Mid$(X$, 63, 4))
     PREUNI = CVD(Mid$(X$, 67, 8))
     CANSTO = CVS(Mid$(X$, 51, 4))
     CANCOM = CVS(Mid$(X$, 55, 4))
     '
     PORDESCUIN = CVS(Mid$(X$, 75, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then PORDESCUIN = 0
     '
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, CODEXT$(CI%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 15)
     If CI% <= 0 Then
       DESCRILI$ = TRIM$(REGLEIDO$("!1OCOMDET", U&))
       Call REPLA(TTXX$, DESCRILI$, 17, 15)
     End If
     Call REPLA(TTXX$, UNISTO$, 34, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANSTO, "F8.1"), 38, 8)
     Call REPLA(TTXX$, UNICOX$, 48, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANCOM, "F9.1"), 52, 9)
       FMTX$ = "F10.4": If PREUNI > 99999 Then FMTX$ = "F10.2"
     Call REPLA(TTXX$, FORMATNUM$(PREUNI, FMTX$), 61, 10)
     Call REPLA(TTXX$, FORMATNUM$(PORDESCUIN, "F5.1"), 72, 5)
     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 78, 8)
     List1.AddItem TTXX$
     '
     'If CAN > 999999 Then
     '  Call MENSERR(24, "Cantidad Excesiva - Max 999999")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     'If CAN < 0.5 Then
     '  Call MENSERR(24, "Falta Cantidad")
     '  Call REPLA(X$, MKS$(0), 47, 4)
     '  Call GRAREG("!CARDETPE", X$, U&)
     '  GoTo 10
     'End If
     '
     '
   Next U&
   'Call CIERRARCH("!CARDETPE")
   'Call CALTOPED
   List1.Refresh
   Screen.MousePointer = 1
End Sub

Sub BUSCAULTIPRE(NPRO%, CIIX%, UNIMAS$, UPREDES$)
   '
   If MODOULPRE% = 2 Then ' costo de reposición
     UPREDES$ = String$(20, 0)
     Call REPLA(UPREDES$, AJUSTI$(UNIMED$(CIIX%), 4), 1, 4)
     Call REPLA(UPREDES$, MKS$(1), 5, 4)
     Call REPLA(UPREDES$, MKD$(COSUNI(CIIX%) * TICAMONE(MONECOSTO%(CIIX%))), 9, 8)
     Call REPLA(UPREDES$, MKS$(0), 17, 4)
     Exit Sub
   End If
   '
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CIIX%))
   FULTICO% = 0
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           If TRIM$(Mid$(YX$, 11, 4)) = TRIM$(UNIMAS$) Then
             If CVI(Mid$(YX$, 5, 2)) >= FULTICO% Then
               FULTICO% = CVI(Mid$(YX$, 5, 2))
               UPREDES$ = Mid$(YX$, 59, 20)
               PREUNANT# = CVD(Mid$(UPREDES$, 9, 8))
               MONEMIXA% = Asc(Mid$(YX$, 122, 1))
               If MONEMIXA% > 1 Then
                 PREUNANT# = PREUNANT# * TICAMONE(MONEMIXA%)
                 Call REPLA(UPREDES$, MKD$(PREUNANT#), 9, 8)
               End If
             End If
           End If
         End If
       End If
     Next U&
   End If
   '
End Sub

Function ULTICOM$(NPRO%, CIIX%)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   FULCO% = 0: UCO$ = REGBLAN$("OCOMDETA")
   '
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CIIX%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           PREUNRE# = CVD(Mid$(YX$, 67, 8))
           If PREUNRE# > 0.00005 Or FULCO% = 0 Then
             If CVI(Mid$(YX$, 5, 2)) >= FULCO% Then
               UCO$ = YX$
               FULCO% = CVI(Mid$(YX$, 5, 2))
             End If
           End If
         End If
       End If
     Next U&
   End If
   '
   ULTICOM$ = UCO$
   '
End Function

Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
    '19/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '19/03/07 (FIN)
'5   For kki% = 0 To 5
'      EDITFORM.Picture1(kki%).Cls
'      EDITFORM.Picture1(kki%).DrawWidth = 3
'      EDITFORM.Picture1(kki%).DrawStyle = 0
'      EDITFORM.Picture1(kki%).FillStyle = 1
'    Next kki%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For kki% = 0 To 5
'      EDITFORM.Picture1(kki%).Cls
'      EDITFORM.Picture1(kki%).Top = 0
'      EDITFORM.Picture1(kki%).Refresh
'      EDITFORM.Picture1(kki%).Height = 7070
'      EDITFORM.Picture1(kki%).Width = 11700
'      EDITFORM.Picture1(kki%).DrawWidth = 3
'      EDITFORM.Picture1(kki%).DrawStyle = 0
'      EDITFORM.Picture1(kki%).FillStyle = 1
'    Next kki%
'    ALTUPAGINA = 7070
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For kki% = 0 To 5
'       EDITFORM.Picture1(kki%).Visible = False
'    Next kki%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'    EDITFORM.HScroll1.Min = 0
'    EDITFORM.HScroll1.Value = 0
'    EDITFORM.HScroll1.Max = 1
'    If ANCHPAGINA > EDITFORM.Frame1.Width Then
'      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
End Sub

Sub PRICOMPRA()
    '
    CPAG% = Val(TEXT21(0))
    FORPAG$ = TRIM$(Eco21(0))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = Val(TEXT21(1))
    EMBALA$ = TRIM$(Eco21(1))
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text13, 52)
    For KKU& = 1 To CARETEX%
      Call REGAPP("!OBSERVOF", RETEX$(KKU&))
    Next KKU&
    Call CIERRARCH("!OBSERVOF")
    '
    PRIPLANOI% = 0
    If CONTROL$("OCOMPRA", "PRIPLANO") = "SI" Then PRIPLANOI% = 1
    PRIESPEC% = 0
    If CONTROL$("OCOMPRA", "PRIESPEC") = "SI" Then PRIESPEC% = 1
    PRIHRINS% = 0
    If CONTROL$("OCOMPRA", "PRIHRINS") = "SI" Then PRIHRINS% = 1
    '
    ' IMPRESION DE LA O/COMPRA
    '
    FORIPRE$ = CONTROL$("", "FORMOCOM")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHEMII% = FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = TRIM$(Text11)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "NUBULTOS"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CODPARAM$(12) = "TRANSPOR"
      CODPARAM$(13) = "POR-DESC"
      CODPARAM$(14) = "IMP-DESC"
      CODPARAM$(15) = "FECH-REC"
      CODPARAM$(16) = "TIPO-VAL"
      CODPARAM$(17) = "IMPO-CHE"
      CODPARAM$(18) = "REF-CONJ"
      CODPARAM$(19) = "DOMI-TRA"
      CODPARAM$(20) = "PERC-IVA"
      '\\\OT-06-0483-RG-14/11/06///
      CODPARAM$(21) = "REF-MAT2"
      CAPARDOC% = 21
      '\\\OT-06-0483-RG-FIN///
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = EMBALA$
      DOCPARAM$(3) = Total(0)
      DOCPARAM$(4) = Text12
      DOCPARAM$(5) = Total(2)
      DOCPARAM$(6) = Total(3)
      DOCPARAM$(7) = "" 'Mid$(X1$, 147, 20)
      DOCPARAM$(8) = "" 'Mid$(X1$, 67, 40)
      DOCPARAM$(9) = "" 'Mid$(X1$, 167, 40)
      DOCPARAM$(10) = "" ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(Val(Total(0)) - Val(Total(1)))
      DOCPARAM$(12) = TRIM$(Text6) ' ATENCION SR.
      DOCPARAM$(13) = TRIM$(PORDESCU)
      DOCPARAM$(14) = Total(1)
      HOII% = HOY(HOS$)
      FECHENT% = FECHANUM(Text14)
      DOCPARAM$(15) = Text14
      If FECHENT% = HOII% Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
        ElseIf FECHENT% < HOII% + 4 Then
          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
      End If
      DOCPARAM$(16) = TRIM$(Label11)
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
      DOCPARAM$(18) = Eco21(1)
      DOCPARAM$(19) = Text10
      DOCPARAM$(20) = Total(4)
      '\\\OT-06-0483-RG-14/11/06///
      DOCPARAM$(21) = Label5.Caption
      '\\\OT-06-0483-RG-FIN///
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CACOLTAB1% = 13
      '
      DESTIDOC% = (-1) * Val(Text1)
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 0
      If Frame16.Visible = False Then
        If CONTROL$("OCOMPRA", "INCLUSPE") = "SI" Then
          INCLUSPE% = 1
        End If
      End If
      '
      REPLAORI% = 0
      If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
         REPLAORI% = 1
      End If
      '
      RECOPROP% = 0
      If CONTROL$("OCOMPRA", "RECOPROP") = "SI" Then
         RECOPROP% = 1
      End If
      '
      REVICODX% = 0
      If CONTROL$("OCOMPRA", "REVICOD") = "SI" Then
         REVICODX% = 1
      End If
      '
      NC1% = DESTIDOC%
      '\\\OT-06-0085-WAR-10/03/06///
      ARCHILIB$ = "!1OCOMDET"
      '\\\OT-06-0085-WAR-FIN///
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PREUNID# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         CPROP$ = CODDX$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         If REVICODX% = 1 Then
            FEREVX = FEREVI%(CI0%)
            RVCDDX$ = TRIM$(REVICOD$(CI0%))
            If RVCDDX$ <> "" Then
              CODDX$ = CODDX$ + " (Letra: " + RVCDDX$ + " - " + FECHATEX$(FEREVX) + ")"
            End If
         End If
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = TRIM$(DESCRIT0$(CI0%))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 DECRIT$ = DECRIT$ + " (" + CPROP$ + ")"
               End If
             End If
           End If
         TETABU1$(2, CAITAB1%) = DECRIT$
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
           KILOX$ = FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
         TETABU1$(11, CAITAB1%) = KILOX$
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         '
         '\\\OT-06-0085-WAR-10/03/06///
         TETABU1$(2, CAITAB1%) = DECRIT$
         If COLTELIB% > 0 Then
           TELIBRE$ = RTrim$(REGLEIDO$(ARCHILIB$, J&))
           If TRIM$(TELIBRE$) = "" Then
             TELIBRE$ = TRIM$(DESCRIT0$(CI0%))
           End If
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 TELIBRE$ = TELIBRE$ + " (" + CPROP$ + ")"
               End If
             End If
           End If
           'FORTELIB$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "DES-MAT/FORMATO")))'\\**//
           FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
           ATELIB% = Val(Mid$(FORTELIB$, 2))
           Call FRATEXTO(TELIBRE$, ATELIB%)
           TETABU1$(2, CAITAB1%) = RETEX$(1)
           For KKU1% = 2 To CARETEX%
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
           Next KKU1%
         End If
         '\\\OT-06-0085-WAR-FIN///
         UNIST$ = Mid$(X$, 11, 4)
         UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
         If TRIM$(Left$(UCase$(UNIST$), 2)) <> UNICO$ Then
           DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + ECOARCH$("UNIMED", UNIST$) + ".)"
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = DEADI$
         End If
         '
         If INCLUSPE% > 0 Then
           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI0%))
           If Len(RNC$) > 0 Then
             INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
             '
             COLUTA% = 0
             URN& = ULTREG&("ESPECIFI")
             '
             For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                 If RENOTA& <= URN& Then
                   TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                   If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
                       CAITAB1% = CAITAB1% + 1
                       TETABU1$(2, CAITAB1%) = TTYY$
                     Else
                       COLUTA% = COLUTA% + 1
                       If COLUTA% > 2 Then COLUTA% = 1
                       If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
                       TETABU1$(COLUTA%, CAITAB1%) = TTYY$
                   End If
                 End If
               End If
             Next U&
             '
           End If
         End If
         '
         If Frame16.Visible = True Then
           CAITAB1% = CAITAB1% + 1
           For KKI% = 0 To CACOLTAB1%
             TETABU1$(KKI%, CAITAB1%) = ""
           Next KKI%
           '
           AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
           ACOLDEI% = Val(Mid$(AA1$, 2))
           '
           Call FRATEXTO(Text28, ACOLDEI%)
           For KKJ% = 1 To CARETEX%
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = RETEX$(KKJ%)
           Next KKJ%
           '
         End If
         '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "ARCON") > 0 Then
         Call CIERRARCH("*.*")
         Call BAJALDISCO
         ANEXOCOM.Show 1
         GoTo 199
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         If CI0% > 0 Then
           If CI0% <= NUMAS% Then
             If PRIPLANOI% = 1 Then Call PRIPLANO(CI0%)
             If PRIESPEC% = 1 Then Call PRIHOJESP(CI0%)
             If PRIHRINS% = 1 Then Call PRIHOJRINSP(CI0%)
           End If
         End If
      Next J&
      '
    End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If EMICONSI% = 1 Then
      If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
        NPRO% = Val(Text1)
        Call ENTRECONSI(NPRO%)
      End If
    End If
    Call CIERRARCH("*.*")
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '

End Sub

Private Sub UcoMat_Click()
   OPCICOMP.Command20_Click
End Sub

Private Sub UcoPro_Click()
   OPCICOMP.Command12_Click
End Sub

Private Sub Ver_Ocom_Click()
   Call Command4_Click
End Sub

Sub GRAOCSIM()
   '
   Static SIMAT%
   '
   If SIMAT% = 0 Then
     SIMAT% = (-1)
     If EXISTE%(LUDAT$ + "BDSIMAT.RFS") > 0 Then
       If EXISTE%(LUDAT$ + "BDSIMAT.DAT") > 0 Then
         SIMAT% = 1
       End If
     End If
   End If
   '
   If SIMAT% > 0 Then
     For U& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", U&)
       CICOM% = CVI(Mid$(X$, 9, 2))  ' Codigo segun lista de materiales a comprar
       CANCOM = CVS(Mid$(X$, 55, 4)) ' Cantidad correspondiente al codigo CI%
       '
       RFF$ = RECFILT$("BDSIMAT", 10, MKI$(CICOM%))
       For UX& = 1 To Len(RFF$) Step 4
         U1& = CVS(Mid$(RFF$, UX&, 4)) '  ULTREG&("BDSIMAT")
         ZZ$ = REGLEIDO$("BDSIMAT", U1&)
         CANSIM = CVS(Mid$(ZZ$, 71, 4)) ' Cantidad correspondiente al codigo CII%
         If CANCOM >= CANSIM Then ' Solo pasa si la cantidad en la OC es >= que la pedida en la SIM
           NROCO& = CVS(Mid$(ZZ$, 131, 4))
           If NROCO& < 1 Then
             NUOCOMAC& = Val(Text11)
             Call REPLA(ZZ$, MKS$(NUOCOMAC&), 131, 4)
               NPROVAS% = Val(Text1)
             Call REPLA(ZZ$, MKI$(NPROVAS%), 129, 2)
               FEMI% = FECHANUM(Text9)
             Call REPLA(ZZ$, MKI$(FEMI%), 135, 2)
               PUNIT = CVD(Mid$(X$, 79, 8))
             Call REPLA(ZZ$, MKS$(PUNIT), 137, 4)
               MONECOSTEO% = CVI(Mid$(X$, 121, 2))
             Call REPLA(ZZ$, MKI$(MONECOSTEO%), 141, 2)
               FENTREOCOM% = CVI(Mid$(X$, 95, 2))
             Call REPLA(ZZ$, MKI$(FENTREOCOM%), 151, 2)
             Call GRAREG("BDSIMAT", ZZ$, U1&)
             CANCOM = CANCOM - CANSIM
           End If
         End If
       Next UX&
     Next U&
   End If
   '
End Sub
