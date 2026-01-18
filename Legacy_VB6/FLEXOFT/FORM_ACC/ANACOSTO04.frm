VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ANACOSTO04 
   BackColor       =   &H00D0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos Standard - Análisis y Presupuestacion de Productos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.TextBox Text45 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4215
      MaxLength       =   25
      TabIndex        =   276
      Top             =   210
      Width           =   2985
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0C0&
      Caption         =   "Acc.Directos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Index           =   1
      Left            =   7350
      TabIndex        =   271
      Top             =   105
      Width           =   1540
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
         Height          =   680
         Left            =   140
         Picture         =   "ANACOSTO04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Planera Electrónica"
         Top             =   315
         Width           =   630
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
         Height          =   680
         Left            =   770
         Picture         =   "ANACOSTO04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   272
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   8310
      Left            =   11025
      ScaleHeight     =   8250
      ScaleWidth      =   1005
      TabIndex        =   260
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox Picture16 
         Height          =   8300
         Left            =   0
         ScaleHeight     =   8235
         ScaleWidth      =   1035
         TabIndex        =   280
         Top             =   0
         Visible         =   0   'False
         Width           =   1100
         Begin VB.Image Image3 
            Height          =   8400
            Left            =   0
            Picture         =   "ANACOSTO04.frx":0614
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
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
         Height          =   640
         Left            =   105
         Picture         =   "ANACOSTO04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   274
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   269
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   2580
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":2C40
         Style           =   1  'Graphical
         TabIndex        =   268
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1890
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Subst"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":2D8A
         Style           =   1  'Graphical
         TabIndex        =   267
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   3255
         Width           =   650
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":2ED4
         Style           =   1  'Graphical
         TabIndex        =   266
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   3960
         Width           =   650
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":31DE
         Style           =   1  'Graphical
         TabIndex        =   265
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
         Width           =   650
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":34E8
         Style           =   1  'Graphical
         TabIndex        =   264
         ToolTipText     =   "Imprime Ficha del Articulo"
         Top             =   6525
         Width           =   650
      End
      Begin VB.CommandButton Command38 
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":3B52
         Style           =   1  'Graphical
         TabIndex        =   263
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   5835
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "ANACOSTO04.frx":3E5C
         Style           =   1  'Graphical
         TabIndex        =   262
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5145
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "ANACOSTO04.frx":429E
         Style           =   1  'Graphical
         TabIndex        =   261
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Index           =   0
         Left            =   -315
         Picture         =   "ANACOSTO04.frx":45A8
         Top             =   7455
         Width           =   1515
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1040
      Left            =   9030
      ScaleHeight     =   975
      ScaleWidth      =   60
      TabIndex        =   41
      Top             =   200
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   5
      Top             =   210
      Width           =   2040
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
      Height          =   330
      Left            =   3285
      Picture         =   "ANACOSTO04.frx":64CA
      TabIndex        =   4
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0C0&
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
      Height          =   1120
      Index           =   0
      Left            =   9300
      TabIndex        =   1
      Top             =   105
      Width           =   1540
      Begin VB.CommandButton Command2 
         Caption         =   "Rest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   770
         Picture         =   "ANACOSTO04.frx":67D4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton Command1 
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
         Height          =   680
         Left            =   140
         Picture         =   "ANACOSTO04.frx":6ADE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   630
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   150
      TabIndex        =   0
      Top             =   1450
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   11324
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12648384
      TabCaption(0)   =   "Insumos"
      TabPicture(0)   =   "ANACOSTO04.frx":6DE8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Frame22"
      Tab(0).Control(3)=   "MARCOFICHA(0)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Componentes"
      TabPicture(1)   =   "ANACOSTO04.frx":6E04
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Elaboración"
      TabPicture(2)   =   "ANACOSTO04.frx":6E20
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Costos"
      TabPicture(3)   =   "ANACOSTO04.frx":6E3C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame16"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame16 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   5895
         Left            =   -74790
         TabIndex        =   118
         Top             =   420
         Width           =   10350
         Begin VB.Frame Frame26 
            Caption         =   "Cobertura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2745
            Left            =   0
            TabIndex        =   186
            Top             =   3150
            Width           =   2325
            Begin VB.TextBox Text44 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   315
               TabIndex        =   241
               Top             =   640
               Width           =   540
            End
            Begin VB.Label Label118 
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
               Left            =   315
               TabIndex        =   255
               Top             =   1980
               Width           =   795
            End
            Begin VB.Label Label117 
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
               Left            =   1420
               TabIndex        =   193
               Top             =   1980
               Width           =   795
            End
            Begin VB.Label Label113 
               Caption         =   "Mg.Bto(%):    Mg.Neto(%):   "
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   105
               TabIndex        =   254
               Top             =   1780
               Width           =   2100
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   247
               Top             =   965
               Width           =   1275
            End
            Begin VB.Label Label89 
               BackStyle       =   0  'Transparent
               Caption         =   "D:"
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
               Index           =   5
               Left            =   105
               TabIndex        =   246
               Top             =   730
               Width           =   2010
            End
            Begin VB.Line Line6 
               BorderColor     =   &H80000005&
               X1              =   0
               X2              =   2310
               Y1              =   1695
               Y2              =   1695
            End
            Begin VB.Line Line5 
               BorderColor     =   &H80000003&
               X1              =   0
               X2              =   2310
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "I.C.Pr:"
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
               Index           =   4
               Left            =   -420
               TabIndex        =   195
               Top             =   2400
               Width           =   1170
            End
            Begin VB.Label Label119 
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
               Left            =   840
               TabIndex        =   194
               Top             =   2325
               Width           =   795
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Margen:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   -1260
               TabIndex        =   192
               Top             =   1365
               Width           =   2115
            End
            Begin VB.Label Label116 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   191
               Top             =   1290
               Width           =   1275
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pr.Neto:"
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
               Index           =   1
               Left            =   -1260
               TabIndex        =   190
               Top             =   1070
               Width           =   2115
            End
            Begin VB.Label Label115 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   189
               Top             =   640
               Width           =   1275
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   188
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pr.Lista:"
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
               Index           =   0
               Left            =   -1260
               TabIndex        =   187
               Top             =   420
               Width           =   2115
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Ultimo Cálculo Integral"
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
            Index           =   1
            Left            =   6770
            TabIndex        =   155
            Top             =   2100
            Width           =   3540
            Begin VB.CommandButton Command55 
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
               Height          =   630
               Left            =   2835
               Picture         =   "ANACOSTO04.frx":6E58
               Style           =   1  'Graphical
               TabIndex        =   156
               ToolTipText     =   "Re-Cálculo General de Costos"
               Top             =   210
               Width           =   540
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "El"
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
               Index           =   33
               Left            =   0
               TabIndex        =   200
               Top             =   600
               Width           =   330
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
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
               Height          =   225
               Index           =   35
               Left            =   1540
               TabIndex        =   199
               Top             =   480
               Width           =   120
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "hs"
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
               Index           =   34
               Left            =   2415
               TabIndex        =   198
               Top             =   600
               Width           =   330
            End
            Begin VB.Label Label95 
               Alignment       =   2  'Center
               BackColor       =   &H80000016&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   1680
               TabIndex        =   197
               Top             =   525
               Width           =   750
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               BackColor       =   &H80000016&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   420
               TabIndex        =   196
               Top             =   525
               Width           =   1065
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Reportes"
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
            Index           =   0
            Left            =   6770
            TabIndex        =   148
            Top             =   0
            Width           =   2325
            Begin VB.PictureBox Picture7 
               BackColor       =   &H80000016&
               Height          =   645
               Left            =   210
               ScaleHeight     =   585
               ScaleWidth      =   1845
               TabIndex        =   152
               Top             =   500
               Width           =   1905
               Begin VB.CommandButton Command54 
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
                  Height          =   645
                  Left            =   0
                  Picture         =   "ANACOSTO04.frx":7162
                  Style           =   1  'Graphical
                  TabIndex        =   153
                  Top             =   0
                  Width           =   645
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Análisis de Costos"
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
                  Index           =   37
                  Left            =   840
                  TabIndex        =   154
                  Top             =   105
                  Width           =   855
               End
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H80000016&
               Height          =   645
               Left            =   210
               ScaleHeight     =   585
               ScaleWidth      =   1845
               TabIndex        =   149
               Top             =   1200
               Width           =   1905
               Begin VB.CommandButton Command50 
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
                  Height          =   645
                  Left            =   0
                  Picture         =   "ANACOSTO04.frx":746C
                  Style           =   1  'Graphical
                  TabIndex        =   150
                  Top             =   0
                  Width           =   645
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Listado Resumen"
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
                  Index           =   36
                  Left            =   840
                  TabIndex        =   151
                  Top             =   105
                  Width           =   1170
               End
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Costo de Suministros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1850
            Left            =   0
            TabIndex        =   137
            Top             =   0
            Width           =   6630
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "C.Propios         Recibidos            Totales"
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
               Index           =   20
               Left            =   2835
               TabIndex        =   270
               Top             =   180
               Width           =   3690
            End
            Begin VB.Line Line7 
               BorderColor     =   &H80000005&
               X1              =   0
               X2              =   6600
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               TabIndex        =   147
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materias Primas:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   146
               Top             =   735
               Width           =   2115
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               TabIndex        =   145
               Top             =   1005
               Width           =   1275
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Piezas Compradas:"
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
               Index           =   1
               Left            =   210
               TabIndex        =   144
               Top             =   1110
               Width           =   2115
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               TabIndex        =   143
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label85 
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
               Height          =   330
               Left            =   5145
               TabIndex        =   142
               ToolTipText     =   "Doble-click para Ver Analisis de Costos de Materiales y Servicios"
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3780
               TabIndex        =   141
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3780
               TabIndex        =   140
               Top             =   1005
               Width           =   1275
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3780
               TabIndex        =   139
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Trabajos de Terceros:"
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
               Index           =   2
               Left            =   210
               TabIndex        =   138
               Top             =   1470
               Width           =   2115
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000003&
               X1              =   0
               X2              =   6600
               Y1              =   420
               Y2              =   420
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Costo de Elaboración"
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
            Index           =   4
            Left            =   0
            TabIndex        =   129
            Top             =   1890
            Width           =   6630
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5145
               TabIndex        =   136
               ToolTipText     =   "Doble-click Para Ver Análisis de Costos de Mano de Obra "
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               TabIndex        =   135
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipos de Producción:"
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
               Index           =   4
               Left            =   0
               TabIndex        =   134
               Top             =   795
               Width           =   2325
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mano de Obra Directa:"
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
               Index           =   3
               Left            =   210
               TabIndex        =   133
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               TabIndex        =   132
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3780
               TabIndex        =   131
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3780
               TabIndex        =   130
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Opción"
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
            Left            =   9240
            TabIndex        =   126
            Top             =   0
            Width           =   1065
            Begin VB.CommandButton Command45 
               BackColor       =   &H00E0E0E0&
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
               Height          =   645
               Left            =   230
               Picture         =   "ANACOSTO04.frx":7776
               Style           =   1  'Graphical
               TabIndex        =   127
               ToolTipText     =   "Padron Maestro de Proveedores "
               Top             =   500
               Width           =   645
            End
            Begin VB.CommandButton Command48 
               BackColor       =   &H00E0E0E0&
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
               Height          =   645
               Left            =   230
               Picture         =   "ANACOSTO04.frx":7A80
               Style           =   1  'Graphical
               TabIndex        =   128
               ToolTipText     =   "Listas de Precios de Venta"
               Top             =   1200
               Width           =   645
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Costo Total Calculado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2745
            Left            =   2415
            TabIndex        =   119
            Top             =   3150
            Width           =   7890
            Begin VB.TextBox Text41 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               TabIndex        =   235
               Top             =   690
               Width           =   750
            End
            Begin VB.TextBox Text43 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5460
               TabIndex        =   239
               Top             =   1065
               Width           =   645
            End
            Begin VB.TextBox Text42 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5460
               TabIndex        =   238
               Top             =   690
               Width           =   645
            End
            Begin VB.TextBox Text39 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5460
               TabIndex        =   240
               Top             =   1785
               Width           =   645
            End
            Begin VB.TextBox Text38 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               TabIndex        =   237
               Top             =   1815
               Width           =   750
            End
            Begin VB.TextBox Text36 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               TabIndex        =   236
               Top             =   1065
               Width           =   750
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   245
               Top             =   1440
               Width           =   1275
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Total Bruto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3255
               TabIndex        =   244
               Top             =   1515
               Width           =   3060
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "GG.Elaboración:               %:"
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
               Index           =   23
               Left            =   180
               TabIndex        =   243
               Top             =   1170
               Width           =   2535
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   242
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   185
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   " Costo s/Venta:             %:"
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
               Index           =   31
               Left            =   4050
               TabIndex        =   184
               Top             =   810
               Width           =   2325
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   " Utilidad:             %:"
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
               Index           =   29
               Left            =   4620
               TabIndex        =   183
               Top             =   1920
               Width           =   1905
            End
            Begin VB.Label Label107 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   182
               Top             =   1815
               Width           =   1275
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Standard Total:"
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
               Index           =   32
               Left            =   4200
               TabIndex        =   181
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   180
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   179
               Top             =   1065
               Width           =   1275
            End
            Begin VB.Label Label103 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   178
               Top             =   2250
               Width           =   1275
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   " Impuestos:             %:"
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
               Index           =   30
               Left            =   4410
               TabIndex        =   177
               Top             =   1170
               Width           =   2115
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Precio Básico de Venta:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   28
               Left            =   3360
               TabIndex        =   176
               Top             =   2370
               Width           =   2955
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
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
               Height          =   225
               Index           =   27
               Left            =   2350
               TabIndex        =   175
               Top             =   1935
               Width           =   330
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Standard Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   26
               Left            =   -420
               TabIndex        =   174
               Top             =   2370
               Width           =   3060
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gastos Grales Admin. y Ventas:"
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
               Index           =   25
               Left            =   105
               TabIndex        =   173
               Top             =   1785
               Width           =   1485
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   172
               Top             =   2250
               Width           =   1275
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   171
               Top             =   1815
               Width           =   1275
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   125
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Costos Directos:"
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
               Index           =   21
               Left            =   525
               TabIndex        =   124
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   123
               Top             =   1065
               Width           =   1275
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "GG.Materiales:               %:"
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
               Index           =   22
               Left            =   315
               TabIndex        =   122
               Top             =   810
               Width           =   2430
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Std de Manufactura:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   24
               Left            =   -420
               TabIndex        =   121
               Top             =   1515
               Width           =   3060
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   120
               Top             =   1440
               Width           =   1275
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "3"
         Height          =   5895
         Index           =   3
         Left            =   -74800
         TabIndex        =   68
         Top             =   420
         Width           =   10410
         Begin VB.Frame Frame4 
            Caption         =   "Análisis de Costo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5800
            Left            =   4620
            TabIndex        =   72
            Top             =   50
            Width           =   5685
            Begin VB.Frame Frame25 
               Caption         =   "Costo Tercerizacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1665
               Index           =   1
               Left            =   210
               TabIndex        =   166
               Top             =   4020
               Width           =   5295
               Begin VB.TextBox Text37 
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
                  Left            =   735
                  TabIndex        =   95
                  Top             =   1260
                  Width           =   1410
               End
               Begin VB.ListBox List2 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   510
                  Left            =   105
                  TabIndex        =   167
                  Top             =   315
                  Width           =   5055
               End
               Begin VB.Label Label63 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   4170
                  TabIndex        =   97
                  ToolTipText     =   "Doble-Click para Seleccionar Fecha"
                  Top             =   915
                  Width           =   990
               End
               Begin VB.Label Label48 
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   735
                  TabIndex        =   94
                  ToolTipText     =   "Doble-click para Seleccionar Proveedor"
                  Top             =   915
                  Width           =   2715
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prov:                                               Fecha:"
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
                  Index           =   18
                  Left            =   255
                  TabIndex        =   248
                  Top             =   990
                  Width           =   4110
               End
               Begin VB.Label Label46 
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   2730
                  TabIndex        =   96
                  ToolTipText     =   "Doble-Click para Seleccionar Moneda"
                  Top             =   1260
                  Width           =   2415
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio:                          Mon:"
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
                  Index           =   19
                  Left            =   105
                  TabIndex        =   168
                  Top             =   1335
                  Width           =   2640
               End
            End
            Begin VB.Frame Frame25 
               Caption         =   "Costo Elaboración"
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
               Index           =   0
               Left            =   3305
               TabIndex        =   165
               Top             =   2310
               Width           =   2200
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "$/U 3's:"
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
                  Index           =   17
                  Left            =   -1155
                  TabIndex        =   258
                  Top             =   705
                  Width           =   1905
               End
               Begin VB.Label Label52 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   840
                  TabIndex        =   257
                  Top             =   630
                  Width           =   1200
               End
               Begin VB.Label Label73 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   840
                  TabIndex        =   170
                  Top             =   270
                  Width           =   1200
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "$/Unid:"
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
                  Index           =   16
                  Left            =   -1470
                  TabIndex        =   169
                  Top             =   345
                  Width           =   2220
               End
            End
            Begin VB.CommandButton Command28 
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
               Height          =   330
               Index           =   0
               Left            =   2100
               TabIndex        =   76
               Top             =   1575
               Width           =   175
            End
            Begin VB.TextBox Text5 
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
               Left            =   1260
               TabIndex        =   74
               Top             =   735
               Width           =   4245
            End
            Begin VB.TextBox Text7 
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
               Left            =   4410
               TabIndex        =   73
               Text            =   " "
               Top             =   420
               Width           =   1095
            End
            Begin VB.TextBox Text9 
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
               Index           =   0
               Left            =   1260
               MaxLength       =   6
               TabIndex        =   75
               Text            =   " "
               Top             =   1575
               Width           =   885
            End
            Begin VB.TextBox Text15 
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
               Left            =   1260
               TabIndex        =   80
               Top             =   2625
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Left            =   1785
               TabIndex        =   81
               Top             =   2625
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Left            =   2205
               TabIndex        =   82
               Top             =   2625
               Width           =   930
            End
            Begin VB.TextBox Text15 
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
               Left            =   1260
               TabIndex        =   83
               Top             =   2940
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Left            =   1785
               TabIndex        =   84
               Top             =   2940
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Left            =   2205
               TabIndex        =   85
               Top             =   2940
               Width           =   930
            End
            Begin VB.TextBox Text15 
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
               Left            =   1260
               TabIndex        =   86
               Top             =   3255
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Left            =   1785
               TabIndex        =   87
               Top             =   3255
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Left            =   2205
               TabIndex        =   88
               Top             =   3255
               Width           =   930
            End
            Begin VB.TextBox Text29 
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
               Left            =   5040
               TabIndex        =   93
               Text            =   " "
               Top             =   3570
               Width           =   465
            End
            Begin VB.TextBox Text30 
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
               Left            =   1260
               TabIndex        =   77
               Text            =   " "
               Top             =   1890
               Width           =   570
            End
            Begin VB.TextBox Text33 
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
               Left            =   2415
               TabIndex        =   78
               Text            =   " "
               Top             =   1890
               Width           =   360
            End
            Begin VB.CommandButton Command29 
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
               Left            =   2775
               TabIndex        =   79
               Top             =   1890
               Width           =   175
            End
            Begin VB.TextBox Text15 
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
               Left            =   1260
               TabIndex        =   89
               Top             =   3570
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Left            =   1785
               TabIndex        =   90
               Top             =   3570
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Left            =   2205
               TabIndex        =   91
               Top             =   3570
               Width           =   930
            End
            Begin VB.TextBox Text34 
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
               Left            =   3720
               TabIndex        =   92
               Text            =   " "
               Top             =   3570
               Width           =   600
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3630
               TabIndex        =   278
               ToolTipText     =   "Tiempos por Unidad"
               Top             =   160
               Width           =   225
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00000000&
               FillColor       =   &H0000C000&
               FillStyle       =   0  'Solid
               Height          =   280
               Left            =   3570
               Shape           =   3  'Circle
               Top             =   210
               Width           =   330
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   420
               Picture         =   "ANACOSTO04.frx":7BCA
               ToolTipText     =   "Abrir Tiempos de Preparación"
               Top             =   2310
               Width           =   480
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Operación:                               Lote Standard:"
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
               Index           =   13
               Left            =   255
               TabIndex        =   117
               Top             =   500
               Width           =   4110
            End
            Begin VB.Label Label43 
               BackColor       =   &H00E0E0E0&
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
               Left            =   2415
               TabIndex        =   116
               Top             =   1575
               Width           =   3090
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   115
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "Preparación:         :      :"
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
               Left            =   105
               TabIndex        =   114
               ToolTipText     =   "Doble-Click para Abrir Tiempos de Preparación"
               Top             =   2700
               Width           =   2220
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   " T.Std.Fijo:         :      :"
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
               Left            =   240
               TabIndex        =   113
               Top             =   3015
               Width           =   2220
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "T.Var.Lote:         :      :"
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
               Left            =   220
               TabIndex        =   112
               Top             =   3330
               Width           =   2430
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Operarios:              Cat:"
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
               Index           =   15
               Left            =   280
               TabIndex        =   111
               Top             =   1995
               Width           =   2220
            End
            Begin VB.Label Label51 
               BackColor       =   &H00E0E0E0&
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
               Left            =   3045
               TabIndex        =   110
               Top             =   1890
               Width           =   2460
            End
            Begin VB.Label Label54 
               BackStyle       =   0  'Transparent
               Caption         =   "    H.    M.     Seg."
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
               Left            =   1260
               TabIndex        =   108
               Top             =   2400
               Width           =   1800
            End
            Begin VB.Label Label55 
               BackColor       =   &H00E0E0E0&
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
               Left            =   1260
               TabIndex        =   107
               Top             =   420
               Width           =   570
            End
            Begin VB.Label Label60 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
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
               Left            =   1260
               TabIndex        =   106
               Top             =   1050
               Width           =   885
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:                      Resp:"
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
               Index           =   14
               Left            =   380
               TabIndex        =   105
               Top             =   1120
               Width           =   2955
            End
            Begin VB.Label Label62 
               BackColor       =   &H00E0E0E0&
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
               Left            =   3045
               TabIndex        =   104
               Top             =   1050
               Width           =   2460
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   " Mant.Herr:         :      :                 Ciclo:             % DD:"
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
               Left            =   225
               TabIndex        =   109
               ToolTipText     =   "Doble-Click para Detalle Mantenimiento Herramentales "
               Top             =   3645
               Width           =   4845
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Operaciones Standard"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Left            =   0
            TabIndex        =   70
            Top             =   50
            Width           =   4530
            Begin VB.CommandButton Command32 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   1395
               Style           =   1  'Graphical
               TabIndex        =   100
               ToolTipText     =   "Elimina la Operación Activa"
               Top             =   2520
               Width           =   630
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   98
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   2520
               Width           =   630
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   750
               Style           =   1  'Graphical
               TabIndex        =   99
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
               Top             =   2520
               Width           =   630
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   105
               TabIndex        =   71
               Top             =   420
               Width           =   4320
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   2520
               Width           =   945
            End
            Begin VB.CommandButton Command42 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Listar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   102
               Top             =   2520
               Width           =   945
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2650
            Left            =   0
            TabIndex        =   69
            Top             =   3200
            Width           =   4530
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2220
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   103
               Top             =   315
               Width           =   4320
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "No Usados"
            Height          =   3900
            Left            =   4830
            TabIndex        =   157
            Top             =   105
            Visible         =   0   'False
            Width           =   3270
            Begin VB.TextBox Text26 
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
               Left            =   2625
               TabIndex        =   164
               Top             =   2310
               Width           =   360
            End
            Begin VB.TextBox Text17 
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
               Left            =   2205
               TabIndex        =   163
               Top             =   2310
               Width           =   360
            End
            Begin VB.TextBox Text15 
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
               Left            =   1680
               TabIndex        =   162
               Top             =   2310
               Width           =   465
            End
            Begin VB.TextBox Text11 
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
               Left            =   1050
               TabIndex        =   161
               Top             =   840
               Width           =   1830
            End
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   160
               Top             =   1890
               Width           =   780
            End
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   159
               Top             =   1575
               Width           =   780
            End
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   158
               Top             =   1260
               Width           =   780
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   5895
         Index           =   2
         Left            =   210
         TabIndex        =   43
         Top             =   420
         Width           =   10200
         Begin VB.Frame Frame12 
            Caption         =   "Consultas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1210
            Left            =   8505
            TabIndex        =   251
            Top             =   4515
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Implosión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   6
               Left            =   210
               TabIndex        =   253
               ToolTipText     =   "Consulta de Conjuntos a los que Pertenece el Elemento Activo"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   4
               Left            =   210
               TabIndex        =   252
               ToolTipText     =   "Muestra Estructura Multinivel"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Actualizar Fórmula / Estructura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            Left            =   210
            TabIndex        =   55
            Top             =   3990
            Width           =   8100
            Begin VB.CheckBox Check1 
               Caption         =   "x Millar"
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
               Left            =   150
               TabIndex        =   279
               Top             =   1300
               Width           =   960
            End
            Begin VB.CommandButton Command27 
               Caption         =   "Append"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "ANACOSTO04.frx":7ED4
               TabIndex        =   63
               ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
               Top             =   410
               Width           =   855
            End
            Begin VB.TextBox Text32 
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
               Left            =   4935
               TabIndex        =   62
               Text            =   " "
               Top             =   420
               Width           =   1170
            End
            Begin VB.TextBox Text31 
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
               Left            =   1260
               TabIndex        =   61
               Top             =   420
               Width           =   1800
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Insert"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "ANACOSTO04.frx":81DE
               TabIndex        =   60
               ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
               Top             =   695
               Width           =   855
            End
            Begin VB.TextBox Text27 
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
               Left            =   6510
               TabIndex        =   59
               Text            =   " "
               Top             =   420
               Width           =   330
            End
            Begin VB.CommandButton Command19 
               Caption         =   "Delete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "ANACOSTO04.frx":84E8
               TabIndex        =   58
               ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
               Top             =   970
               Width           =   855
            End
            Begin VB.CommandButton Command31 
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
               Left            =   3045
               TabIndex        =   57
               Top             =   420
               Width           =   175
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "ANACOSTO04.frx":87F2
               TabIndex        =   56
               ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label41 
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
               Height          =   360
               Left            =   3885
               TabIndex        =   64
               Top             =   420
               Width           =   435
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción:"
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
               Left            =   105
               TabIndex        =   67
               Top             =   945
               Width           =   1065
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código:                                      U-M:           Uso:                       D:"
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
               Index           =   11
               Left            =   380
               TabIndex        =   66
               Top             =   525
               Width           =   6105
            End
            Begin VB.Label Label40 
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
               Height          =   570
               Left            =   1260
               TabIndex        =   65
               Top             =   945
               Width           =   5580
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   210
            ScaleHeight     =   375
            ScaleWidth      =   8025
            TabIndex        =   53
            Top             =   390
            Width           =   8085
            Begin VB.Label Label27 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Codigo                   Descripcion                                       U-M              Uso     D"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   105
               Width           =   7995
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   210
            TabIndex        =   52
            Top             =   820
            Width           =   8100
         End
         Begin VB.Frame Frame11 
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   8505
            TabIndex        =   48
            Top             =   2940
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Enlaces"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   0
               Left            =   210
               TabIndex        =   51
               ToolTipText     =   "Listado de Estucturas y Enlaces"
               Top             =   380
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   1
               Left            =   210
               TabIndex        =   50
               ToolTipText     =   "Lista de Materiales Multinivel"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "L.Compras"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   2
               Left            =   210
               TabIndex        =   49
               ToolTipText     =   "Lista de Compras"
               Top             =   1100
               Width           =   1170
            End
         End
         Begin VB.CommandButton Command41 
            Height          =   435
            Left            =   8715
            Picture         =   "ANACOSTO04.frx":8AFC
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   380
            Width           =   1170
         End
         Begin VB.Frame Frame14 
            Caption         =   "Utilidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   8505
            TabIndex        =   46
            Top             =   1785
            Width           =   1590
            Begin VB.CommandButton Command40 
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
               Height          =   280
               Index           =   3
               Left            =   210
               TabIndex        =   250
               ToolTipText     =   "Control de Integridad de la Base de Datos"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   5
               Left            =   210
               TabIndex        =   249
               ToolTipText     =   "Nueva Fórmula por Copia de una Pre-Existente"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame FRAME15 
            Caption         =   "Revisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   8505
            TabIndex        =   44
            Top             =   945
            Width           =   1590
            Begin VB.Label Label59 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   210
               TabIndex        =   45
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Line Line1 
            X1              =   8295
            X2              =   8295
            Y1              =   380
            Y2              =   840
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   5685
         Index           =   0
         Left            =   -74790
         TabIndex        =   12
         Top             =   525
         Width           =   10320
         Begin VB.Frame Frame8 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2430
            Left            =   105
            TabIndex        =   230
            Top             =   3150
            Width           =   4740
            Begin VB.TextBox Text40 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1905
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   231
               Top             =   315
               Width           =   4320
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Reposición"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2430
            Left            =   5040
            TabIndex        =   21
            Top             =   3150
            Width           =   5160
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
               Height          =   285
               Left            =   3720
               TabIndex        =   259
               Top             =   795
               Width           =   1150
            End
            Begin VB.TextBox Text21 
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
               Left            =   3465
               TabIndex        =   39
               Top             =   1575
               Width           =   1410
            End
            Begin VB.TextBox Text13 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               MaxLength       =   2
               TabIndex        =   31
               Top             =   795
               Width           =   360
            End
            Begin VB.CommandButton Command15 
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
               Left            =   1620
               TabIndex        =   30
               Top             =   795
               Width           =   175
            End
            Begin VB.TextBox Text14 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3720
               TabIndex        =   28
               Top             =   420
               Width           =   1005
            End
            Begin VB.CommandButton Command16 
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
               Left            =   4725
               TabIndex        =   27
               Top             =   420
               Width           =   175
            End
            Begin VB.TextBox Text12 
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
               Left            =   1260
               TabIndex        =   25
               Top             =   420
               Width           =   1410
            End
            Begin VB.TextBox Text10 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               MaxLength       =   5
               TabIndex        =   23
               Top             =   1575
               Width           =   780
            End
            Begin VB.CommandButton Command14 
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
               Left            =   2040
               TabIndex        =   22
               Top             =   1575
               Width           =   175
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "$ / Unidad:"
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
               Index           =   10
               Left            =   2100
               TabIndex        =   40
               Top             =   1650
               Width           =   1290
            End
            Begin VB.Label Label35 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               TabIndex        =   34
               Top             =   1170
               Width           =   3615
            End
            Begin VB.Label Label33 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               TabIndex        =   33
               Top             =   1950
               Width           =   3615
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda:                        % G.Nacion:"
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
               Index           =   8
               Left            =   420
               TabIndex        =   32
               Top             =   885
               Width           =   3705
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Costeo:"
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
               Index           =   6
               Left            =   2460
               TabIndex        =   29
               Top             =   315
               Width           =   1185
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Reposición:"
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
               Index           =   7
               Left            =   -105
               TabIndex        =   26
               Top             =   315
               Width           =   1290
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedor Habitual:"
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
               Index           =   9
               Left            =   -315
               TabIndex        =   24
               ToolTipText     =   "Doble-Click para Control Cambios Proveedor Habitual"
               Top             =   1470
               Width           =   1500
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Ultimas Compras"
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
            Left            =   5040
            TabIndex        =   232
            Top             =   105
            Width           =   5160
            Begin VB.ListBox List4 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   210
               TabIndex        =   233
               ToolTipText     =   "Doble-Click para Registrar Precio de Compra como Costo de Reposición"
               Top             =   685
               Width           =   4740
            End
            Begin VB.Label Label92 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha     Proveedor           Precio  "
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   210
               TabIndex        =   234
               Top             =   420
               Width           =   4740
            End
            Begin VB.Line Line4 
               X1              =   210
               X2              =   4960
               Y1              =   2785
               Y2              =   2785
            End
            Begin VB.Line Line3 
               X1              =   4965
               X2              =   4965
               Y1              =   400
               Y2              =   2800
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Características"
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
            Left            =   105
            TabIndex        =   13
            Top             =   105
            Width           =   4740
            Begin VB.TextBox Text23 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1155
               TabIndex        =   227
               Top             =   1470
               Width           =   3375
            End
            Begin VB.TextBox Text22 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1155
               TabIndex        =   226
               Top             =   1890
               Width           =   3375
            End
            Begin VB.TextBox Text19 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1155
               TabIndex        =   216
               Top             =   2415
               Width           =   1305
            End
            Begin VB.TextBox Text24 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3210
               TabIndex        =   20
               Top             =   2415
               Width           =   1305
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
               Height          =   330
               Left            =   1700
               TabIndex        =   17
               Top             =   945
               Width           =   175
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   16
               Text            =   " "
               Top             =   945
               Width           =   570
            End
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
               Height          =   330
               Left            =   1700
               TabIndex        =   15
               Top             =   525
               Width           =   175
            End
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   14
               Text            =   " "
               Top             =   525
               Width           =   570
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "         Color:"
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
               Index           =   3
               Left            =   105
               TabIndex        =   229
               Top             =   1805
               Width           =   990
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Talle ó Medida:"
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
               Index           =   2
               Left            =   105
               TabIndex        =   228
               Top             =   1410
               Width           =   990
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Económico:"
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
               Index           =   5
               Left            =   -105
               TabIndex        =   217
               Top             =   2350
               Width           =   1200
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Material:"
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
               Left            =   -315
               TabIndex        =   38
               Top             =   450
               Width           =   1380
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unidad de Medida:"
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
               Left            =   -105
               TabIndex        =   37
               Top             =   890
               Width           =   1185
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kilos/ Unid:"
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
               Index           =   4
               Left            =   2310
               TabIndex        =   36
               Top             =   2350
               Width           =   870
            End
            Begin VB.Label Label30 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1995
               TabIndex        =   19
               Top             =   945
               Width           =   2535
            End
            Begin VB.Label Label29 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1995
               TabIndex        =   18
               Top             =   525
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "No Usados"
         Height          =   4215
         Left            =   -74475
         TabIndex        =   201
         Top             =   630
         Width           =   4320
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
            Height          =   330
            Left            =   1810
            TabIndex        =   221
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3675
            TabIndex        =   220
            Top             =   945
            Width           =   1305
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   219
            Text            =   " "
            Top             =   105
            Width           =   360
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   525
            TabIndex        =   218
            Top             =   945
            Width           =   1305
         End
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
            Height          =   330
            Left            =   1500
            TabIndex        =   213
            Top             =   1785
            Width           =   175
         End
         Begin VB.TextBox Text25 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1155
            TabIndex        =   212
            Text            =   " "
            Top             =   1785
            Width           =   360
         End
         Begin VB.CommandButton Command34 
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
            Height          =   330
            Left            =   5010
            TabIndex        =   211
            Top             =   1200
            Width           =   175
         End
         Begin VB.CommandButton Command35 
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
            Height          =   330
            Left            =   5010
            TabIndex        =   210
            Top             =   1710
            Width           =   175
         End
         Begin VB.CommandButton Command36 
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
            Height          =   330
            Left            =   1485
            TabIndex        =   207
            Top             =   735
            Width           =   175
         End
         Begin VB.TextBox Text35 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   945
            MaxLength       =   3
            TabIndex        =   206
            Text            =   "  "
            Top             =   735
            Width           =   570
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   945
            MaxLength       =   3
            TabIndex        =   203
            Text            =   "  "
            Top             =   315
            Width           =   570
         End
         Begin VB.CommandButton Command13 
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
            Height          =   330
            Left            =   1485
            TabIndex        =   202
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote Máximo:"
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
            Left            =   3465
            TabIndex        =   225
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Minimo:"
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
            Left            =   210
            TabIndex        =   224
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Criterio de Reposicion:"
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
            Left            =   210
            TabIndex        =   223
            Top             =   0
            Width           =   1380
         End
         Begin VB.Label Label32 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2100
            TabIndex        =   222
            Top             =   105
            Width           =   2900
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F.Des:"
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
            Left            =   210
            TabIndex        =   215
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label Label39 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1155
            TabIndex        =   214
            Top             =   2205
            Width           =   3375
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo I.V.A.:"
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
            Left            =   -105
            TabIndex        =   209
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label57 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1785
            TabIndex        =   208
            Top             =   735
            Width           =   2535
         End
         Begin VB.Label Label31 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1785
            TabIndex        =   205
            Top             =   315
            Width           =   2535
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Contable:"
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
            Left            =   -420
            TabIndex        =   204
            Top             =   210
            Width           =   1290
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   10
         Top             =   4305
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   9
         Top             =   3360
         Width           =   15
      End
   End
   Begin VB.Image ImageExcel 
      Enabled         =   0   'False
      Height          =   330
      Left            =   6720
      Picture         =   "ANACOSTO04.frx":8E06
      Stretch         =   -1  'True
      Top             =   210
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Pl.Nro:"
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
      Index           =   41
      Left            =   3360
      TabIndex        =   277
      Top             =   290
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M::"
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
      Index           =   40
      Left            =   5985
      TabIndex        =   275
      Top             =   680
      Width           =   795
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   11865
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   105
      TabIndex        =   256
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6615
      TabIndex        =   35
      Top             =   925
      Width           =   580
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   11
      Top             =   925
      Width           =   750
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Index           =   38
      Left            =   -210
      TabIndex        =   8
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Index           =   39
      Left            =   400
      TabIndex        =   7
      Top             =   290
      Width           =   795
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abri 
         Caption         =   "Abrir"
      End
      Begin VB.Menu phl5 
         Caption         =   "-"
      End
      Begin VB.Menu Anuevo 
         Caption         =   "Nuevo Item de Stock"
      End
      Begin VB.Menu Acopia 
         Caption         =   "Alta por Copia"
      End
      Begin VB.Menu plh0 
         Caption         =   "-"
      End
      Begin VB.Menu Guar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Imp 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Sal 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu DesCam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu phl4 
         Caption         =   "-"
      End
      Begin VB.Menu SustCod 
         Caption         =   "Sustitución de Código de Articulo"
      End
      Begin VB.Menu BaLog 
         Caption         =   "Baja Lógica Individual"
      End
   End
   Begin VB.Menu ListCon 
      Caption         =   "Listados y Consultas"
      Begin VB.Menu GralMa 
         Caption         =   "Generales del Maestro"
         Begin VB.Menu AtGral 
            Caption         =   "Atributos Generales"
         End
         Begin VB.Menu StoLot 
            Caption         =   "Stock Minino y Lote de Reposición"
         End
         Begin VB.Menu CapInv 
            Caption         =   "Captura para Inventarios"
         End
      End
      Begin VB.Menu CosRep 
         Caption         =   "Costos de Reposición"
         Begin VB.Menu NumPro 
            Caption         =   "Por Numero de Proveedor"
         End
         Begin VB.Menu CodArt 
            Caption         =   "Por Código de Articulo"
         End
      End
      Begin VB.Menu ItInac 
         Caption         =   "Items Inactivos"
         Begin VB.Menu ConRes 
            Caption         =   "Consulta y Restauración"
         End
         Begin VB.Menu Listad 
            Caption         =   "Listado"
         End
      End
      Begin VB.Menu phl6 
         Caption         =   "-"
      End
      Begin VB.Menu CosRan 
         Caption         =   "Costos Calculados por Rango de Código"
      End
   End
   Begin VB.Menu AccDire 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaePro 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu LisPre 
         Caption         =   "Lista de Precios de Venta"
      End
      Begin VB.Menu php6a 
         Caption         =   "-"
      End
      Begin VB.Menu Otro 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu ConGen 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu phl8 
         Caption         =   "-"
      End
      Begin VB.Menu CalCos 
         Caption         =   "Re-Calculo General de Costos"
      End
   End
   Begin VB.Menu ayu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ANACOSTO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****/////COMANDO OCULTO/////*****
'Se oculto en el Picture1, el Command25
'("Control de Integridad de Base de Datos")
'*****/////**************/////*****
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%(15), MODIFIC%(15)
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
'
Dim RESECO$(128), DESOPE$(128)
'
Dim PROVINT%(32767)
Dim FOCOREP%
'
Dim PORGAGE(7, 255), MODOCA%(7, 255)
Dim PREVENETO#, MODOCALCU%(7)
Dim CAMBIOEXCEL%
Dim xlBook As Object
Sub ValidaExcel()
 NN% = EXISTE%(LUDAT$ + "\COSTINDU\" + ATRIB1$(CI%) + ".xls")
 NN% = NN% + EXISTE%(LUDAT$ + "\COSTINDU\" + ATRIB1$(CI%) + "_1.xls")
 If NN% > 0 Then
   ImageExcel.Visible = True
   ImageExcel.Enabled = True
 Else
   ImageExcel.Visible = False
   ImageExcel.Enabled = False
 End If
End Sub
'
Private Sub IMPORMASVIE()
   
   Command6.Enabled = False
   Call ABREMASTER
   On Error Resume Next
   Master.MoveLast
   While Master.RecordCount > 0
     Master.MoveLast
     Master.Delete
   Wend
   '
   Call ABRESTRUC
   On Error Resume Next
   Formulas.MoveLast
   While Formulas.RecordCount > 0
     Formulas.MoveLast
     Formulas.Delete
   Wend
   '
   Call ABRESECOP
   On Error Resume Next
   Secoper.MoveLast
   While Secoper.RecordCount > 0
     Secoper.MoveLast
     Secoper.Delete
   Wend
   '
   FIN& = ULTREG&("MASTER")
   For U& = 2 To FIN&
     '
     REG& = U&   'CIXI% + 1
     REMAS$ = REGLEIDO$("MASTER", REG&)
     CIXI% = CVI(Left$(REMAS$, 2))
     '
     CODIG$ = TRIM$(Mid$(REMAS$, 77, 16))
     STAT% = 1
       If Mid$(CODIG$, 16, 1) = Chr$(96) Then STAT% = (-1)
     CODIG$ = TRIM$(Left$(CODIG$, 15))   ' para eliminar caracter 96
     Descrip$ = TRIM$(Mid$(REMAS$, 5, 64))
       If Descrip$ = "" Then Descrip$ = " "
     UNIME$ = TRIM$(Mid$(REMAS$, 109, 2))
     TIMA% = Asc(Mid$(REMAS$, 93, 1))
     CRIT% = Asc(Mid$(REMAS$, 94, 1))
     MONEI% = Asc(Mid$(REMAS$, 95, 1))
     COSUN = CVS(Mid$(REMAS$, 111, 4))
     FEX = CVI(Mid$(REMAS$, 115, 2)):   FECOS$ = TRIM$(FECHATEX$(FEX))
     PORNA = CVI(Mid$(REMAS$, 127, 2)) / 100
     GRUCO% = CVI(Mid$(REMAS$, 125, 2))
     GRUIVA% = Asc(Mid$(REMAS$, 3, 1))
     NPRO% = CVI(Mid$(REMAS$, 123, 2))
     STOMI = CVI(Mid$(REMAS$, 105, 2)): If STOMI < 0 Then STOMI = 100 * Abs(STOMI)
     LOSTAN = CVI(Mid$(REMAS$, 107, 2)): If LOSTAN < 0 Then LOSTAN = 100 * Abs(LOSTAN)
     LOMAX = CVI(Mid$(REMAS$, 103, 2)): If LOMAX < 0 Then LOMAX = 100 * Abs(LOMAX)
         CC% = Asc(Mid$(REMAS$, 96, 1))
     Colors$ = LCOLOR$(CC%)
         CT% = Asc(Mid$(REMAS$, 97, 1))
     TALLES$ = LTALLE$(CT%)
     PESUN = CVS(Mid$(REMAS$, 99, 4))
         'FEX = FEREVI%(CI%)
         'Text25.TEXT = FECHATEX$(FEX)
         'CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
         'Text25.TEXT = TRIM$(Str$(CCOMPO%))
         'Label39.Caption = TRIM$(ECOARCH$("TACOMPO", Chr$(CCOMPO%)))
         'Text36 = REVICOD$(CI%)
         '
     With Master
45      If CIXI% > NUMAS% + 1 Then
          .AddNew
          !CODINT = 1 + RecordCount
          .Update
          GoTo 45
        End If
        '
        On Error GoTo 0
        'On Error Resume Next
        .AddNew
        !CODIGO = CODIG
        !CODINT = CIXI%
        !Status = STAT%
        !DESCRIPCION = Descrip
        !UMEDIDA = UNIME
        !TIMATE = TIMA%
        !CRITREP = CRIT%
        !MONECOS = MONEI%
        !COSUNI = COSUN
        On Error Resume Next
        !FECOSTO = FECOS
        On Error GoTo 0
        !PORNACIO = PORNA
        !PROVHABI = NPRO%
        !STOMIN = STOMI
        !STOMAX = LOMAX
        !LOREPOS = LOREPO
        !PESUNI = PESUN
        .Update
        .Bookmark = .LastModified
        UPDATMASTER% = 1
    End With
    On Error GoTo 0
  Next U&
  '
  Master.Close
  '
  ' .Append CREACAMPOS.CreateField("Color", dbText)
  ' .Append CREACAMPOS.CreateField("Talle", dbText)
  ' .Append CREACAMPOS.CreateField("Descompo", dbInteger)
  ' .Append CREACAMPOS.CreateField("Grupiva", dbInteger)
  ' .Append CREACAMPOS.CreateField("Grucoven", dbInteger)
  ' .Append CREACAMPOS.CreateField("Observa", dbMemo)
  ' .Append CREACAMPOS.CreateField("Especificacion", dbMemo)
  '
  FIN& = ULTREG("ESTRUNUE")
  For REG& = 1 To FIN&
    X$ = REGLEIDO$("ESTRUNUE", REG&)
    CICONJ% = CVI(Left$(X$, 2))
    DPS$ = Mid$(X$, 3, 120)
    REVIFORMU = CVI(Mid$(X$, 123, 2))
    If ULTREG&("ESTRUDOS") >= REG& Then
       DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
    End If
    '
    If CICONJ% > 0 Then
      RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CICONJ%))
      If Len(RFF$) >= 4 Then
        RELA& = CVS(RFF$)
        If RELA& > 0 Then
          If RELA& <= ULTREG&("ESTRULAR") Then
            DPS$ = DPS$ + Mid$(REGLEIDO$("ESTRULAR", RELA&), 3)
          End If
        End If
      End If
      '
      While Len(DPS$) Mod 8 <> 0
        DPS$ = DPS$ + Chr$(0)
      Wend
      '
      For K% = 1 To Len(DPS$) Step 8
        DPSX$ = Mid$(DPS$, K%, 8)
        CIXXII% = CVI(Left$(DPSX$, 2))
        If CIXXII% > 0 Then
          If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
            With Formulas
              On Error Resume Next
              .AddNew
              !CODICONJ = CICONJ%
              !CODXCONJ = CODEXT$(CICONJ%)
              !CODIELEM = CIXXII%
              !CODXELEM = CODEXT$(CIXXII%)
              !USOBRUTO = CVS(Mid$(DPSX$, 3, 4))
              !PLADISPO = Asc(Mid$(DPSX$, 7, 1))
              .Update
              .Bookmark = .LastModified
            End With
          End If
        End If
      Next K%
    End If
  Next REG&
  Formulas.Close
  '
  Call GENECOMAS
  '
  With Secoper
    FIN& = ULTREG("SECOPER")
    For REG& = 1 To FIN&
      X$ = REGLEIDO$("SECOPER", REG&)
      '
      .AddNew
         CI% = CVI(Left$(X$, 2))
         RESECOX$ = X$
      !CODICONJ = CI%
      !CODXCONJ = CODEXT$(CI%)
      !NumeOper = CVI(Mid$(RESECOX$, 3, 2))
      !DescOper = TRIM$(Mid$(RESECOX$, 5, 48))
      !statoper = 1
      !CodMaq0 = TRIM$(Mid$(RESECOX$, 53, 6))
      !CodMaq1 = TRIM$(Mid$(RESECOX$, 103, 6))
      !CodMaq2 = TRIM$(Mid$(RESECOX$, 109, 6))
      !CodMaq3 = TRIM$(Mid$(RESECOX$, 115, 6))
      !HorsPrep = CVS(Mid$(RESECOX$, 59, 4))
      !HorsFijo = CVS(Mid$(RESECOX$, 63, 4))
      !MinuStan = CVS(Mid$(RESECOX$, 67, 4))
      !HorsLimp = CVS(Mid$(RESECOX$, 79, 4))
      !HorsMHer = CVS(Mid$(RESECOX$, 95, 4))
      !PieCiclo = CVS(Mid$(RESECOX$, 99, 4))
      !PorDemor = CVS(Mid$(RESECOX$, 75, 4))
      If !MinuStan > 0.005 Then
         !Cadencia = 60 / !MinuStan
      End If
      !CoefSolp = CVS(Mid$(RESECOX$, 87, 4))
      !LOTESTAN = CVS(Mid$(RESECOX$, 91, 4))
      !CantiOps = CVS(Mid$(RESECOX$, 83, 4))
      !ProgCNum = TRIM$(Mid$(RESECOX$, 121, 8))
      '    !HoProFab = TRIM$(Mid$(RESECOX$, 129, 16))
      '    !ProvOper = Str$(CVI(Mid$(RESECOX$, 129, 2)))
      '    !PreUnid = CVD(Mid$(RESECOX$, 131, 8))
      '    !MonePrec = CVI(Mid$(RESECOX$, 139, 2))
      '    !FechPrec = CVDAT(CVI(Mid$(RESECOX$, 141, 2)))
      .Update
      '
    Next REG&
    .Close
  End With
  '
  MsgBox "FIN IMPORTACION"
  Command6.Enabled = True
  '
End Sub

Private Sub Abri_Click()
Call Command9_Click
End Sub

Private Sub Acopia_Click()
   Call Command22_Click
End Sub

Private Sub Anuevo_Click()
   Call Command21_Click
End Sub



Private Sub AtGral_Click()
Call OPLIMAS04.Command9_Click
End Sub

Private Sub ayu_Click()
Call Command4_Click
End Sub

Private Sub BaLog_Click()
Call Command24_Click
End Sub

Private Sub CalCos_Click()
Call Command55_Click
End Sub

Private Sub CapInv_Click()
Call OPLIMAS04.COMMAND15_CLICK
End Sub

Private Sub Check1_Click()
   If NOCHECK% < 1 Then
     If Check1.Value = 1 Then
         If PORMILLAR% < 1 Then
           If HABIPORMI% < 1 Then
                Check1.Value = 0
                PORMILLAR% = 0
                Exit Sub
             Else
                For KU% = 1 To List1.ListCount
                  XXX$ = List1.List(KU% - 1)
                  List1.RemoveItem KU% - 1
                  USOX$ = FORMATNUM$(1000 * XVALO(Mid$(XXX$, 49, 12)), "F12.4")
                  Call REPLA(XXX$, USOX$, 49, 12)
                  List1.AddItem XXX$, KU% - 1
                Next KU%
                PORMILLAR% = 1
           End If
         End If
       ElseIf Check1.Value = 0 Then
         If PORMILLAR% = 1 Then
           For KU% = 1 To List1.ListCount
             XXX$ = List1.List(KU% - 1)
             List1.RemoveItem KU% - 1
             USOX$ = FORMATNUM$(XVALO(Mid$(XXX$, 49, 12)) / 1000, "F12.4")
             Call REPLA(XXX$, USOX$, 49, 12)
             List1.AddItem XXX$, KU% - 1
           Next KU%
           PORMILLAR% = 0
         End If
     End If
   End If
   NOCHECK% = 0
End Sub

Private Sub CodArt_Click()
Call OPLIMAS04.Command4_Click
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   OPT% = 0
10 OPT% = ALTERNA%("Actualización de Costos por Grupo Contable\Otros Accesos Directos")
   '
   If OPT% = 1 Then
      Call OPLIMAS04.ACTULISCOSGCONT
      GoTo 10
   ElseIf OPT% = 2 Then
      Call ACCDIR("ANACOS")
       GoTo 10
   End If
   '
   Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("PLANEL04/" + TRIM$(Text1), "Planera Electrónica")
   Command7.Enabled = True
End Sub

Private Sub COMMLAB_Change()
   If COMMLAB = "PRIMOD" Then
      COMMLAB.Caption = ""
      Call PRIANAMODI
    ElseIf COMMLAB = "PRISUMI" Then
      COMMLAB.Caption = ""
      Call PRIANASUMI
   End If
End Sub

Private Sub ConGen_Click()
Call Command37_Click
End Sub

Private Sub ConInt_Click()
Call Command25_Click
End Sub

Private Sub ConRes_Click()
Call OPLIMAS04.Command3_Click
End Sub

Private Sub CosRan_Click()
Call Command50_Click
End Sub

Private Sub DesCam_Click()
Call Command2_Click
End Sub

Private Sub Form_Load()
    '
    SSTab1.Tab = 3
    Call CENTRAFORM(Me)
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1(4).BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5(0).BackColor = &HFCD7B6: Me.Frame5(1).BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.Visible = False
      Me.Frame12.Visible = False
      
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      Me.FRAME15.BackColor = &HFCD7B6
      Me.Frame16.BackColor = &HFCD7B6
      Me.Frame17.BackColor = &HFCD7B6
      Me.Frame18.BackColor = &HFCD7B6
      Me.Frame19.BackColor = &HFCD7B6
      Me.Frame20(0).BackColor = &HFCD7B6: Me.Frame20(1).BackColor = &HFCD7B6
      
      Me.Frame22.BackColor = &HFCD7B6
      Me.Frame23.BackColor = &HFCD7B6
      Me.Frame24.BackColor = &HFCD7B6
      Me.Frame25(0).BackColor = &HFCD7B6: Me.Frame25(1).BackColor = &HFCD7B6
      Me.Frame26.BackColor = &HFCD7B6
      Me.MARCOFICHA(0).BackColor = &HFCD7B6: Me.MARCOFICHA(2).BackColor = &HFCD7B6: Me.MARCOFICHA(3).BackColor = &HFCD7B6
      
      Me.Label113(5).BackColor = &HFCD7B6
      Me.Label92.BackColor = &HFCD7B6
      Me.Label27.BackColor = &HFCD7B6
      Me.Picture5.BackColor = &HFCD7B6
      
      Me.Picture5.BackColor = &HFCD7B6
      
      Me.Check1.BackColor = &HFCD7B6
      Me.SSTab1.BackColor = &HFCD7B6
    End If

    Show
    Refresh
    ' Call VERJOBID
    '
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(UCase$(EPAC$), "ESCORIAL") > 0 Then
       Text45.FontSize = 8
       Text45.Width = 1980
       Text45.FontSize = 11
    End If
    '
    Command45.Enabled = False
    If EXISTE%("ARCHIG04.EXE") > 0 Then Command45.Enabled = True
    '
    Command48.Enabled = False
    If EXISTE%("LISPRE04.EXE") > 0 Then Command48.Enabled = True
    '
    Command7.Enabled = False
    If EXISTE%("PLANEL04.EXE") Then Command7.Enabled = True
    '
    HABIPORMI% = 0
    If CONTROL$("AMASTO", "PORMILES") = "SI" Then
      HABIPORMI% = 1
    End If
    '
    Call CARTAVALISTO
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "DECICAL"))
    If DCC$ = "" Then
        DECICALCOS% = 4
        Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
      Else
        DECICALCOS% = Val(DCC$)
    End If
    If DECICALCOS% < 0 Or DECICALCOS% > 5 Then DECICALCOS% = 4
    '
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(1, KU&) = GGG: MODOCA%(1, KU&) = 1: Next
      Text41 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(2, KU&) = GGG: MODOCA%(2, KU&) = 1: Next
      Text36 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENADM"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(3, KU&) = GGG: MODOCA%(3, KU&) = 1: Next
      Text38 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORCOMIS"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(4, KU&) = GGG: MODOCA%(4, KU&) = 1: Next
      Text42 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(5, KU&) = GGG: MODOCA%(5, KU&) = 1: Next
      Text43 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(6, KU&) = GGG: MODOCA%(6, KU&) = 1: Next
      Text39 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORDESCU"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(7, KU&) = GGG: MODOCA%(7, KU&) = 1: Next
      Text44 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL("ANACOSTO", "FECALCOS"))
      Label34(4) = DCC$
    DCC$ = TRIM$(CONTROL("ANACOSTO", "HOCALCOS"))
      Label95 = DCC$
      '
      If EXISTE%(LUDAT$ + "GAGENCOS.DAT") > 0 Then
         For KU& = 1 To ULTREG&("GAGENCOS")
           XXX$ = REGLEIDO$("GAGENCOS", KU&)
           GRGR% = CVI(Left$(XXX$, 2))
           If GRGR% >= 0 Then
             If GRGR% <= 255 Then
               NORI% = CVI(Mid$(XXX$, 3, 2))
               If NORI% >= 0 Then
                 If NORI% <= 7 Then
                   PORGAGE(NORI%, GRGR%) = CVS(Mid$(XXX$, 37, 4))
                   MODOCA%(NORI%, GRGR%) = CVI(Mid$(XXX$, 41, 2))
                 End If
               End If
             End If
           End If
         Next KU&
      End If
      '
    Label10(2) = CONTROL$("AMASTO", "ATRIPAR1")
      If Label10(2) = "" Then Label10(2) = ". Atributo 'A':"
      If Left$(Label10(2), 1) = "." Then Label10(2) = " " + Mid$(Label10(2), 2)
    Label10(3) = CONTROL$("AMASTO", "ATRIPAR2")
      If Label10(3) = "" Then Label10(3) = ". Atributo 'B':"
      If Left$(Label10(3), 1) = "." Then Label10(3) = " " + Mid$(Label10(3), 2)
      '
    Call ACVEPROINT
    '
    Text1.SetFocus
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Resize()
   IMPLOSI.Hide
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MODOEDIT% = 1 Then
     Cancel = 1
     Call MENSAEDIT
     Exit Sub
   End If
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command1_Click()
   command1.Enabled = False
   If MODOEDIT% = 1 Then Call MENSAEDIT: GoTo 99
   GRABATODO% = 1
   Call GRADATITEM
   Call ValidaExcel
   Call GENECOMAS
   GRABATODO% = 0
99 command1.Enabled = True
End Sub

Private Sub Command10_Click()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Command11_Click()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Command12_Click()
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Command13_Click()
   Call SELECHO("GRUCOVEN")
   VALONUE$ = VALACT1$("GRUCOVEN")
   If VALONUE$ <> "" Then
      Text16.TEXT = VALONUE$
   End If
End Sub

Private Sub Command14_Click()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text10.TEXT))))
   End If
End Sub

Private Sub COMMAND15_CLICK()
   Call SELECHO("TAMONED")
   VALONUE$ = VALACT1$("TAMONED")
   If VALONUE$ <> "" Then
      Text13.TEXT = VALONUE$
   End If
End Sub

Private Sub Command16_Click()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text14.TEXT = VALONUE$
   End If
End Sub

Private Sub Command18_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Edición no Válida")
       Exit Sub
    End If
    '
    Text31.TEXT = Left$(List1.TEXT, 16)
    CDYY% = CODINT%(Text31.TEXT)
    Label40.Caption = DESCRIT0$(CDYY%)
    Label41.Caption = UNIMED$(CDYY%)
    Text32.TEXT = TRIM$(Mid$(List1.TEXT, 49, 12))
    Text27.TEXT = TRIM$(Mid$(List1.TEXT, 61, 4))
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command18.Enabled = False
    Command19.Enabled = False
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    Text32.SetFocus
    MODOEDIT% = 1
    MODIFIC%(2) = 1
    '
End Sub

Private Sub Command19_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Borrado no Válida")
       Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    MODIFIC%(2) = 1
    '
End Sub

Private Sub Command2_Click()
   '
   If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
      Exit Sub
   End If
   '
   For KKII% = 0 To 15
     YACAR%(KKII%) = 0
     MODIFIC%(KKII%) = 0
   Next KKII%
   '
   Command18.Enabled = True
   Command19.Enabled = True
   MODOEDIT% = 0
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Command20_Click()
   Call SELECHO("TACOMPO")
   VALONUE$ = VALACT1$("TACOMPO")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Command21_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command21.Enabled = False
   Call GRADATITEM
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
      CODINUE$ = Mid$(VDEF$, 77, 16)
      Text1.TEXT = CODINUE$
      SSTab1.Tab = 0
   End If
   Call GENECOMAS
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   Dim FIELDVAL(256)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   CII% = CODINT%(Text1.TEXT)
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   'If REMAS& < 2 Or REMAS& > ULTREG&("MASTER") Then
   '   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   '   Command25.SetFocus
   '   GoTo 99
   'End If
   '
   VDEF$ = REGBLAN$("MASTER")
   CINUE% = 1 + NUMAS%
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
   Call REPLA(VDEF$, UNIMED$(CII%), 109, 2)
   Call REPLA(VDEF$, Chr$(TIMATE%(CII%)), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   Call ABREMASTER
   If OKX% = 1 Then
     CACAM% = Articulos.TableDefs("Master").Fields.Count
     Master.FindFirst ("CODINT = " & CII%)
     If Master.NoMatch = False Then
       Master.Edit
       For KU% = 1 To CACAM%
         FIELDVAL(KU%) = Master.Fields(KU% - 1).Value
       Next KU%
       '
       Master.FindFirst ("CODINT = " & CINUE%)
       If Master.NoMatch = False Then
         Master.Edit
         For KU% = 1 To CACAM%
           NOCAMPO$ = TRIM$(UCase$(Master.Fields(KU% - 1).Name))
           If NOCAMPO$ <> "CODIGO" Then
             If NOCAMPO$ <> "DESCRIPCION" Then
               If NOCAMPO$ <> "CODINT" Then
                 If NOCAMPO$ <> "ALIAS" Then
                   If NOCAMPO$ <> "UMEDIDA" Then
                     Master.Fields(KU% - 1).Value = FIELDVAL(KU%)
                   End If
                 End If
               End If
             End If
           End If
         Next KU%
         Master.Update
       End If
       '
       Call COPYFORMU(CII%, CINUE%)
       Call COPYRUTA(CII%, CINUE%)
     End If
     '
     CODINUE$ = Mid$(VDEF$, 77, 16)
     Text1.TEXT = CODINUE$
     SSTab1.Tab = 0
     '
   End If
   '
99 Call GENECOMAS
   Command22.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Command23.Enabled = False
   If VALCODART% < 1 Then GoTo 99
   '
   VDEF$ = ""
10 OBX$ = OBJPLA$("ENTRACODNUE", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   CODINUE$ = TRIM$(OBX$)
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
     Master.FindFirst ("CODINT = " & CI%)
     Master.Edit
     Master!CODIGO = CODINUE$
     Master.Update
     UPDATMASTER% = 1
   End If
   '
   Text1.TEXT = CODINUE$
   '
99 Call GENECOMAS
   Screen.MousePointer = 1
   Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command24.Enabled = False
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), 15)
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   CII% = CODINT%(Text1.TEXT)
   If CII% > 0 Then
     Master.FindFirst ("CODINT = " & CI%)
     Master.Edit
     Master!Status = (-1)
     Master.Update
     UPDATMASTER% = 1
     Call BLANFORMU
     Text1 = ""
   End If
   '
99 Call GENECOMAS
   '
   TTXX$ = "Baja Lógica del Código '" + TRIM$(COD$) + "' Completada."
   TTXX$ = TTXX$ + "\  \¿ Desea Eliminar el Artículo\de Todas las Listas de Precios ?\  \(Esta Operación no Podrá Revertirse)"
   If COMALTER%(TTXX$ + "\\No, Continuar\Si, Eliminarlo") = 2 Then
      NULIS& = ULTREG&("LISTAS")
      For U& = 1 To NULIS&
         Screen.MousePointer = 11
         LISNO% = U&
         ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                          ALISTA$ = "0" + ALISTA$
                                        Wend
         APRECIO$ = "LIPRE" + ALISTA$
         If ULTREG&(APRECIO$) > 0 Then
            Call FILTERDELETE(APRECIO$, 1, MKI$(CII%))
         End If
      Next U&
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
    Command25.Enabled = False
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
   '
   If VALCODART% < 1 Then Exit Sub
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(Val(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
   CAIT% = CAIT% + 1
   ILISTA& = List1.ListIndex
   If ILISTA& < 0 Then ILISTA& = 0
   '
   List1.AddItem XXX$, ILISTA&
   List1.ListIndex = ILISTA&
   List1.Refresh
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text27.TEXT = ""
   Text31.SetFocus
   MODIFIC%(2) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
   If VALCODART% < 1 Then Exit Sub
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(Val(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
   CAIT% = CAIT% + 1
   '
   List1.AddItem XXX$
   List1.ListIndex = List1.ListCount - 1
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text27.TEXT = ""
   Text31.SetFocus
   MODIFIC%(2) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command28_Click(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command29_Click()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If MODOEDIT% = 1 Then Call MENSAEDIT: GoTo 99
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call FINAL("")
99 Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command30.Enabled = False
   Call GRADATITEM
   For KKII% = 0 To 15: YACAR%(KKII%) = 0: Next KKII%
   Call COPYRUTAS
   Command30.Enabled = True
End Sub

Private Sub COMMAND31_CLICK()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
    '
End Sub

Private Sub Command32_Click()
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   ILISTA& = List3.ListIndex
   If ILISTA& >= 0 Then
     List3.RemoveItem (ILISTA&)
     For UI& = ILISTA& To List3.ListCount - 1
       RESECO$(UI&) = RESECO$(UI& + 1)
       DESOPE$(UI&) = DESOPE$(UI& + 1)
     Next UI&
   End If
   '
   Label55.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Label51.Caption = ""
   MODIFIC(3) = 1
   '
   If List3.ListCount > 0 Then
     List3.ListIndex = 0
   End If
   '
End Sub

Private Sub Command33_Click()
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   NOPY% = 10
   If List3.ListCount > 0 Then
     ULI& = List3.ListCount
     NOPY% = 10 * (1 + Int(Val(Left$(List3.List(ULI& - 1), 9) / 10)))
   End If
   RESECO$(List3.ListCount) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(List3.ListCount) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Agregada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(3) = 1
   List3.AddItem SECO$
   List3.ListIndex = List3.ListCount - 1
   Text5.SetFocus
End Sub

Private Sub Command34_Click()
   Call SELECHO("TATALLES")
   TALLENUE$ = VALACT1$("TATALLES")
   If TALLENUE$ <> "" Then
      Text23.TEXT = TALLENUE$
   End If
End Sub

Private Sub Command35_Click()
   Call SELECHO("TACOLOR")
   COLONUE$ = VALACT1$("TACOLOR")
   If COLONUE$ <> "" Then
      Text22.TEXT = COLONUE$
   End If
End Sub

Private Sub Command36_Click()
   Call SELECHO("GRUPIVA")
   VALONUE$ = VALACT1$("GRUPIVA")
   If VALONUE$ <> "" Then
      Text35.TEXT = VALONUE$
   End If
End Sub

Private Sub Command37_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   For KKI% = 0 To 15: YACAR%(KKI%) = 0: Next KKI%
   '
   If DERACCE%("STCOSTO", "Configuración Procesos de Costos") >= 2 Then
     OpAnCoAcc.Show 1
     Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación.")
     Call CONECRUN("FSETUP04/ANACOSTO", "")
     Call FINAL("")
   End If
   '
   Call ACVEPROINT
   Call CARDATITEM
   '
   CIXI% = CODINT%(Text1)
   MONEI% = Val(Text13.TEXT)
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIXI%) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + PORNACIO(CI%) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   If SSTab1.Tab = 2 Then
     If List3.ListIndex >= 0 Then
       VDEF$ = RESECO$(List3.ListIndex)
       Mid$(VDEF$, 151, 8) = MKD(Val(Text37))
       RESECO$(List3.ListIndex) = VDEF$
       MONEXI% = Val(Mid$(Label46, 49, 6))
       TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
       Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
       MODIFIC%(3) = 1
     End If
   End If
   '
End Sub

Private Sub Command38_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command38.Enabled = False
   OPLIMAS04.Show 1
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   If SSTab1.Tab <> 3 Then
        PRINTFORM
      Else
        Call PRIFICHCOS
   End If
   Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call HELPONLINE("ANACOS04")
   Command4.Enabled = True
End Sub

Private Sub Command40_Click(Index As Integer)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   Command40(Index).Enabled = False
   '
   Call GRADATITEM
   Select Case Index
      Case 0
        Call LIESTRU
      Case 1
        Call LIMULTI
      Case 2
        Call LIMULIN
      Case 3
        Call REDESPIE
      Case 4
        Call COMULTI
      Case 5
        Call COPYSTRUC
      Case 6
        Call LIMPLOSI
   End Select
   '
   Command40(Index).Enabled = True
   '
End Sub

Private Sub Command41_Click()
  '
  If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
  Call GRADATITEM
  '
  If IUCOD% > 0 Then
    IUCOD% = IUCOD% - 1
    Text1.TEXT = UCOD$(IUCOD% + 1)
  End If
End Sub

Private Sub Command42_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command42.Enabled = False
   Call GRADATITEM
   Call LISTRUTAS
   Command42.Enabled = True
End Sub

Private Sub Command45_Click()
   '
   Command45.Enabled = False
   If EXISTE%("ARCHIG04.EXE") > 0 Then
       Call CONECRUN("ARCHIG04/AMAPRO", "Padron Maestro de Proveedores")
     Else
       Call OPNOIN("")
   End If
   Command45.Enabled = True
   '
End Sub

Private Sub Command48_Click()
   '
   Command48.Enabled = False
   Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
   Command48.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   ILISTAX& = List3.ListIndex
   If ILISTAX& < 0 Or ILISTAX& >= List3.ListCount Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   NOPEP% = Val(Left$(List3.List(ILISTAX&), 4))
   NOPEA% = 0
   If ILISTAX& > 0 Then
     NOPEA% = Val(Left$(List3.List(ILISTAX& - 1), 4))
   End If
   NOPY% = (NOPEA% + NOPEP%) / 2
   If NOPY% <= NOPEA% Or NOPY% >= NOPEP% Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   For UI& = List3.ListCount To ILISTAX& Step -1
     If UI& > 0 Then
       RESECO$(UI&) = RESECO$(UI& - 1)
       DESOPE$(UI&) = DESOPE$(UI& - 1)
     End If
   Next UI&
   '
   RESECO$(ILISTAX&) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(ILISTAX&) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Insertada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(3) = 1
   List3.AddItem SECO$, ILISTAX&

   List3.ListIndex = ILISTAX&
   Text5.SetFocus
   '
End Sub

Private Sub Command51_Click()

End Sub

Private Sub Command50_Click()
   Command50.Enabled = False
   If ULTREG&("RECOSCAL") < 1 Then
       Call MENSERR(24, "Falta Generar Archivo Resumen")
     Else
       FILICOS04.Show 1
   End If
   Command50.Enabled = True
End Sub

Private Sub Command54_Click()
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
   Command54.Enabled = False
   Call PRIFICHCOS
   Call PRIANASUMI
   Call PRIANAMODI
   Command54.Enabled = True
   '
End Sub

Private Sub Command55_Click()
   Command55.Enabled = False
   RECALCOSTO% = 1
   '
   On Error Resume Next
   RECALCOS04.Show
   On Error GoTo 0
   Command55.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Guar_Click()
Call Command1_Click
End Sub
Private Sub ImageExcel_DblClick()
    '
    ' PRIMERO DEBE GUARDAR
    If CAMBIOEXCEL% > 0 Then
       Call MENSERR(24, "Debe guardar los cambios primero")
       Exit Sub
    End If
    '
    OPX% = COMALTER%("Opciones Alternativas\\Planilla de Costos\Lista de Componentes Valorizados")
    '
    CI% = CODINT%(Text1)
    If CI% < 1 Then Exit Sub
    '
    Mycodigo = ATRIB1$(CI%)
    If OPX% = 2 Then Mycodigo = ATRIB1$(CI%) + "_1"
    'MyLugar = (LUDAT$ + "COSTINDU\" + MyCodigo + ".xls")
    Mytabla = (LUDAT$ + "BDAccess\Articulos.MDB")
    '
    On Error Resume Next
    Set xlBook = CreateObject("excel.application")
    xlBook.Workbooks.Open FileName:=(LUDAT$ + "COSTINDU\" + Mycodigo + ".xls")
    xlBook.ActiveSheet.Cells(1, 20) = Mytabla '(Es T1 en Excel)
    xlBook.Visible = True
    'Set xlBook = GetObject(LUDAT$ + "COSTINDU\" + Mycodigo + ".xls")
       
       If Err.Number = 287 Then
          Exit Sub
       ElseIf Err.Number > 0 Then
         Call MENSERR(24, "Se produjo Error " & Err.Number)
         GoTo 99
       End If
       
       'xlBook.ActiveSheet.MyLugar = MyLugar
       'xlBook.ActiveSheet.MyTabla = MyTabla
       'xlBook.application.Visible = True
       'xlBook.Windows(1).Visible = True
       
99 Set xlBook = Nothing

End Sub

Private Sub Imp_Click()
Call Command39_Click
End Sub

Private Sub Label1_Click()
   Text7 = 1
End Sub

Private Sub Label1_DblClick()
   CIXI% = CODINT%(Text1)
   If CIXI% > 0 Then
     Text7 = TRIM$(Str$(LOTESTAN(CIXI%)))
   End If
End Sub

Private Sub Label103_Change()
   If MODOCALCU%(4) = 1 Then Label110 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text42) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(5) = 1 Then Label104 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text43) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(6) = 1 Then Label107 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text39) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label104_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label105_Change()
   PREVEN# = 0
   If (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
   '
   Call Text42_Change
   Call Text43_Change
   Call Text39_Change
End Sub

Private Sub Label110_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label111_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) - Val(Label111)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label112_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
   '
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
   Label118 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * (Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - (Val(Label85) + Val(Label93))) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label118 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label118.ForeColor = RGB(0, 0, 0)
   If Val(Label118) < 0 Then Label118.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label114_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Text44) / 100), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label115_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label116_Change()
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
   Label118 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * (Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - (Val(Label85) + Val(Label93))) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label118 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label118.ForeColor = RGB(0, 0, 0)
   If Val(Label118) < 0 Then Label118.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label117_Change()
   Label119 = ""
   If Val(Text39) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Text39)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label13_DblClick()
   CI% = CODINT%(Text1)
   If CI% > 0 Then
     Call TRALOFIL("RECAMPH", 1, MKI$(CI%))
     If ULTREG&("!RECAMPH") < 1 Then
          Call MENSERR(24, "Sin Cambios Registrados\para Código " + TRIM$(CODEXT$(CI%)))
        Else
          VTB% = VENTABU%("RECAMPH/TITUPAN=Cambios Proveedor Habitual Código " + TRIM$(CODEXT$(CI%)))
     End If
   End If
End Sub

Private Sub Label37_Click()
    Call ABREMASTER
    Set MASTERECO = Articulos.OpenRecordset("SELECT * FROM Master ORDER BY CODINT ASC;", dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    MASTERECO.MoveLast
    If Err Then Exit Sub
    '
    CAREC% = MASTERECO.RecordCount
    If CAREC% < 1 Then Exit Sub
    '
    MASTERECO.MoveFirst
    JJ& = 1
10  MASTERECO.MoveNext
    JJ& = JJ& + 1
    If MASTERECO!CODINT <> JJ& Then
       AAA = BBB
    End If
    GoTo 10
    '

End Sub

Private Sub Label38_DblClick()
   Screen.MousePointer = 11
   INI& = ULTREG&("MOVISTO") - 8192
   If INI& < 1 Then INI& = 1
   JJ& = 0
   Call COPYSTRU("MOVISTO", "MOVISTF")
   For U& = INI& To ULTREG&("MOVISTO")
     XX$ = REGLEIDO$("MOVISTO", U&)
     JJ& = JJ& + 1
     Call GRAREG("MOVISTF", XX$, JJ&)
   Next U&
   Call SETULTREG("MOVISTF", JJ&)
   Call CIERRARCH("MOVISTF")
   Call CIERRARCH("MOVISTO")
   Screen.MousePointer = 1
End Sub

Private Sub Label46_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 159, 2) = MKI(Val(Mid$(Label46, 49, 6)))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = Val(Mid$(Label46, 49, 6))
   TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   MODIFIC%(3) = 1
End Sub

Private Sub Label46_DblClick()
   '
   Call SELECHO("TAMONED")
   MONEI% = Val(TRIM$(VALACT1$("TAMONED")))
   If MONEI% > 0 Then
     MONEX$ = AJUSTI$(VALACT2$("TAMONED"), 48) + Str$(MONEI%)
     Label46.Caption = MONEX$
   End If
   '
End Sub

Private Sub Label47_DblClick(Index As Integer)
   If Index = 0 Then
     If Val(Label55) > 0 Then
       VDEF$ = Space$(64) + String$(64, 0)
       Call REPLA(VDEF$, Text1, 1, 16)
       Call REPLA(VDEF$, Text3, 17, 48)
       Call REPLA(VDEF$, MKI$(Val(Label55)), 65, 2)
       Call REPLA(VDEF$, Text5, 67, 30)
       '
       CIXI% = CODINT%(Text1): NOPYI% = Val(Label55)
       TPXX = Val(Text15(0)) + Val(Text17(0)) / 60 + Val(Text26(0)) / 3600
       CAOPRE = CANOPREP(CIXI%, NOPYI%)
       CTOPRE% = CATOPREP(CIXI%, NOPYI%)
       Call REPLA(VDEF$, MKS$(TPXX), 97, 4)
       Call REPLA(VDEF$, MKS$(CAOPRE), 101, 4)
       Call REPLA(VDEF$, MKI$(CTOPRE%), 105, 2)
       OBX$ = OBJPLA$("CARHORPREP", VDEF$)
       '
       If OBX$ <> "" Then
         Screen.MousePointer = 11
         CAOPRE = CVS(Mid$(OBX$, 101, 4))
         CTOPRE% = CVI(Mid$(OBX$, 105, 2))
         If ECOARCH$("CATOPER", Chr$(CTOPRE%)) = "" Then
           CTOPRE% = 0
         End If
         '
         Call ABRESECOP
         With Secoper
           .FindFirst ("CodIConj = " & CIXI% & "AND NumeOper = " & NOPYI% & " AND StatOper > 0")
           If .NoMatch = False Then
               .Edit
               !NUOPER_PREP = CAOPRE
               !CATOPER_PREP = CTOPRE%
               .Update
              Else
               Call MENSERR(24, "Imposible Grabar")
           End If
         End With
         '
         Text5 = Mid$(OBX$, 67, 30)
         TSTT$ = HORMINSEDIEM$(60 * 60 * CVS(Mid$(OBX$, 97, 4)))
         Text15(0).TEXT = Left$(TSTT$, 3)
         Text17(0).TEXT = Mid$(TSTT$, 5, 2)
         Text26(0).TEXT = Mid$(TSTT$, 8, 7)
         '
         Call CALCOSOPE
         Screen.MousePointer = 1
         '
       End If
     End If
   End If
   '
   'NUEVO PARA MANTENIMIENTO HERRAMENTAL
   If Index = 4 Then
     If Val(Text1) > 0 Then
       NOPE% = Val(List3.TEXT)
       If NOPE% > 0 Then
         CLAVE$ = MKI$(CI%) + MKI$(NOPE%)
         Call FILTERFILE("MANTHERR", "", 1, CLAVE$)
         VTB% = VENTABU%("MANTHERR")
         If VTB% >= 0 Then
           TOTMINU = 0
           For II& = 1 To ULTREG("!MANTHERR")
             X$ = REGLEIDO("!MANTHERR", II&)
             MINUL = CVS(Mid$(X$, 33, 4))
             CAPIE = CVS(Mid$(X$, 37, 4))
             If CAPIE > 0 Then
                TOTMINU = TOTMINU + MINUL / CAPIE
              End If
           Next II&
           '
           TOTMINU = TOTMINU * 100
           Call FILTERUPDATE("MANTHERR", "", 1, CLAVE$)
           YYY$ = HORMINSEDIEM$(TOTMINU * 60)
           Text15(4) = Mid$(YYY$, 1, 3)
           Text17(4) = Mid$(YYY$, 5, 2)
           Text26(4) = Mid$(YYY$, 8, 7)
           Text34 = "100.0"
         End If
       End If
     End If
   End If
End Sub

Private Sub Label48_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 149, 2) = MKI(Val(Mid$(Label48, 49, 6)))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Label48_DblClick()
   '
   Call SELECHO("ACREDOR")
   NPRO% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NPRO% > 0 Then
     PROVEX$ = AJUSTI$(VALACT2$("ACREDOR"), 48) + Str$(NPRO%)
     Label48.Caption = PROVEX$
   End If
   '
End Sub

Private Sub Label63_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 161, 2) = MKI(FECHANUM(Label63))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Label63_DblClick()
   Call SELECHO("SELFECHA")
   FECHU$ = TRIM$(VALACT1$("SELFECHA"))
   If FECHU$ <> "" Then
     Label63.Caption = FECHU$
   End If
End Sub

Private Sub Label65_Change()
   'Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Text38) / 100), 3, 12, DECICALCOS%, 2))
   Call Text38_Change
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label68_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label70_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label77_Change()
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label80_Change()
   'Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Text36) / 100), 3, 12, DECICALCOS%, 2))
   Call Text36_Change
   Label68.Refresh
End Sub

Private Sub Label80_DBLClick()
   Label80.Enabled = False
   If ANACOSMOD.List1.ListCount > 0 Then
     ANACOSMOD.Show 1
   End If
   Label80.Enabled = True
End Sub

Private Sub Label85_Change()
   'Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Text41) / 100), 3, 12, DECICALCOS%, 2))
   Call Text41_Change
End Sub

Private Sub Label85_dblClick()
   Label85.Enabled = False
   If ANACOSMAT.List1.ListCount > 0 Then
     ANACOSMAT.Show 1
   End If
   Label85.Enabled = True
End Sub

Private Sub Label93_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label96_Change()
   Label105 = Label96
End Sub

Private Sub LisPre_Click()
Call Command48_Click
End Sub

Private Sub List1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   If IUCOD% < 128 Then
     If TRIM$(Text1) <> "" Then
       If Text1 <> UCOD$(IUCOD%) Then
         IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = Text1.TEXT
                              ICOMPO%(IUCOD%) = List1.ListIndex
       End If
     End If
   End If
   '
   Call GRADATITEM
   COD$ = TRIM$(Left$(List1.TEXT, 16))
   If COD$ <> "" Then
     Text1.TEXT = COD$
     Call CARDATITEM
   End If
   '
End Sub

Private Sub List3_Click()
   Call MUEDATOPER
End Sub


Function VALCODART%()
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
End Function
'
Sub BLANFORMU()
   '
   Call BLANFORCOS
   
   Text3.TEXT = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
   Text4.TEXT = ""
   Label30.Caption = ""
   Text6.TEXT = ""
   Label29.Caption = ""
   Text8.TEXT = ""
   Label32.Caption = ""
   Text13.TEXT = ""
   Label35.Caption = ""
   Text16.TEXT = ""
   Label31.Caption = ""
   Text35.TEXT = ""
   Label57.Caption = ""
   Text10.TEXT = ""
   Label33.Caption = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   Text20.TEXT = ""
   Text12.TEXT = ""
   Text14.TEXT = ""
   Text21.TEXT = ""
   Text22.TEXT = ""
   Text23.TEXT = ""
   Text24.TEXT = ""
   Text25.TEXT = ""
   Text40 = ""
   Label39.Caption = ""
   '
   List1.Clear
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   List3.Clear
   Label55.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Label51.Caption = ""
   PREVENETO# = 0
   ImageExcel.Visible = False
   ImageExcel.Enabled = False
   '
End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
         If CI% < 0 Then
             Call MENSERR(24, "Codigo Anulado o de Baja")
           Else
             Call MENSERR(24, "Código Inexistente o no Válido")
         End If
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    Text45 = NUPLANO$(CI%)
    MODIFIC(0) = 0
    FOCOREP% = 1
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENITEM
      Case 1
        Call CARFORMULA
      Case 2
        Call CARHORUTA
      Case 3
        Call CARCOSTO
    End Select
    CAMBIOEXCEL% = 0
    Screen.MousePointer = 1
    '
End Sub

Sub CARGENITEM()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
         UNID$ = UNIMED$(CI%)
      Text4.TEXT = UNID$
      Label30.Caption = ECOARCH$("UMEDIDA", UNID$)
         TIMA% = TIMATE%(CI%)
      Text6.TEXT = TRIM$(Str$(TIMA%))
      Label29.Caption = ECOARCH$("TATIMAT", Chr$(TIMA%))
         GRUCO% = GRUCOVEI%(CI%)
      Text16.TEXT = TRIM$(Str$(GRUCO%))
      Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(GRUCO%))
         GRUIVA% = GRUPIVA%(CI%)
      Text35.TEXT = TRIM$(Str$(GRUIVA%))
      Label57.Caption = ECOARCH$("GRUPIVA", Chr$(GRUIVA%))
      '
      Text12.TEXT = CSTRING$(MKS$(COSUNI(CI%)), 3, 11, 4, 2)
         FEX = FECOSTO%(CI%)
      Text14.TEXT = FECHATEX$(FEX)
         MONEI% = MONECOSTO%(CI%)
      Text13.TEXT = TRIM$(Str$(MONEI%))
      Label35.Caption = ECOARCH$("TAMONED", Chr$(MONEI%))
         PORNA = PORNACIO(CI%)
      Text28.TEXT = AJUSTD$(FORMATNUM$(PORNA, "F7.2"), 8)
      '
      COSUDOL = Val(Text12)
      If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNA / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
      End If
      Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
         '
         NPRO% = PROVHABI%(CI%)
      Text10.TEXT = TRIM$(Str$(NPRO%))
      Label33.Caption = ECOARCH$("ACREDOR", MKI$(NPRO%))
         '
         CRIT% = CRITREP%(CI%)
      Text8.TEXT = TRIM$(Str$(CRIT%))
      Label32.Caption = ECOARCH$("TACRIRE", Chr$(CRIT%))
      Text19.TEXT = CSTRING$(MKS$(LOTESTAN(CI%)), 3, 10, 1, 2)
      Text18.TEXT = CSTRING$(MKS$(STOMIN(CI%)), 3, 10, 1, 2)
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(CI%)), 3, 10, 1, 2)
      '
      Text22.TEXT = COLOR$(CI%)
      Text23.TEXT = TALLE$(CI%)
      Text24.TEXT = LTITPE$(PESUNI(CI%))
         '
         OBSX$ = OBSERVAIT$(CI%)
      Text40 = OBSX$
'GoTo 90
         '
         'ESTO QUE HACE????
'         '
'      CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
'      Text25.TEXT = TRIM$(Str$(CCOMPO%))
'      Label39.Caption = TRIM$(ECOARCH$("TACOMPO", Chr$(CCOMPO%)))
      '
      List4.Clear
      UAGRE$ = ""
      'NUEVO
      '
      Call ABRECOMPRAS
      SQLX$ = "SELECT * FROM OCOMDETA"
      SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
      SQLX$ = SQLX$ + " AND Stat_Ocom >= 0 "
      SQLX$ = SQLX$ + " AND N_OrServ < 1 "
      SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + " ORDER BY Femi_Ocom ASC "
      Set ULTICOMPTMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      With ULTICOMPTMP
        '
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          GoTo 90
        End If
        On Error GoTo 0
        '
        Do While .EOF = False
          Z$ = Space$(64)
          FEX = CVINT(!Femi_Ocom)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          NPXX% = !Npro_Ocom
          RPP$ = ECOARCH$("ACREDOR", MKI$(NPXX%))
          If RPP$ <> "" Then
            Call REPLA(Z$, RPP$, 11, 12)
          End If
          Call REPLA(Z$, !Valo_Unid, 25, 12)
          If Z$ <> UAGRE$ Then '  para no repetir si la O/C tiene varias entregas programadas
            List4.AddItem Z$
            UAGRE$ = Z$
          End If
          .MoveNext
        Loop
      End With
         
'*******///// CODIGO ANTIGUO/////*********
'      URE& = ULTREG&("OCOMDETA")
'      RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
'      For KKI% = 1 To Len(RFF$) Step 4
'        REOCO& = CVS(Mid$(RFF$, KKI%, 4))
'        If REOCO& > 0 Then
'          If REOCO& <= URE& Then
'            XX$ = REGLEIDO$("OCOMDETA", REOCO&)
'            STAT% = Asc(Mid$(XX$, 124, 1))
'            If STAT% <> 9 Then    ' Or CVS(Mid$(XX$, 99, 4)) >= 0.05 Then       ' no esta anulada
'              NUORSE& = CVS(Mid$(XX$, 117, 4))
'              If NUORSE& < 1 Then ' excluye orden de servicios
'                Z$ = Space$(64)
'                FEX = CVI(Mid$(XX$, 5, 2))
'                Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
'                NPXX% = CVI(Mid$(XX$, 7, 2))
'                RPP$ = ECOARCH$("ACREDOR", MKI$(NPXX%))
'                If RPP$ <> "" Then
'                  Call REPLA(Z$, RPP$, 11, 12)
'                End If
'                COEFUNI = CVS(Mid$(XX$, 63, 4))
'                If COEFUNI < 0.0001 Then COEFUNI = 1
'                IUNI# = CVD(Mid$(XX$, 79, 8)) / COEFUNI
'                Call REPLA(Z$, "$", 25, 1)
'                Call REPLA(Z$, CSTRING$(MKD$(IUNI#), 3, 10, 4, 2), 27, 10)
'                If TRIM$(Mid$(XX$, 43, 12)) <> "" Then
'                  Call REPLA(Z$, Mid$(XX$, 43, 12), 25, 12)
'                End If
'                If Z$ <> UAGRE$ Then '  para no repetir si la O/C tiene varias entregas programadas
'                   List4.AddItem Z$
'                   UAGRE$ = Z$
'                End If
'              End If
'            End If
'          End If
'        End If
'      Next KKI%
      '
90    MODIFIC%(0) = 0
      YACAR%(0) = 1
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

Sub CARESPECI()
   '
   If YACAR%(1) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
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
30    Text28.TEXT = Left$(TTYZ$, U&)
      Text28.Refresh
      '
      YACAR%(1) = 1
      MODIFIC%(1) = 0
      '
   End If
   '
End Sub
'
'
Sub CARHORUTA()
   '
   If YACAR%(3) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      List3.Clear
      Label55 = ""
      Text5 = ""
      Text7 = ""
      Text11 = ""
      Label43 = ""
      Text29 = ""
      Text30 = ""
      Text33 = ""
      Text34 = ""
      Label51 = ""
      For III% = 0 To 4
        Text15(III%) = ""
        Text17(III%) = ""
        Text26(III%) = ""
      Next III%
      For III% = 0 To 3
        Text9(III%) = ""
      Next III%
      '
      ULTOP% = 0
      SQLX$ = "SELECT * FROM Secoper"
      SQLX$ = SQLX$ & " WHERE CodIConj = " & CI%
      SQLX$ = SQLX$ & " AND StatOper >0 "
      SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
      Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      REBLA$ = REGBLAN$("SECOPER") + String$(64, 0)
      IOP% = 0
      On Error Resume Next
      With SecopTemp
        .MoveFirst
        If Err Then GoTo 25
        '
        Do Until .EOF = True
          If IsNull(!AltOper) = True Or !AltOper < 1 Then
            NOPY% = !NumeOper
            DOPE$ = !DescOper
            LOTESTA = LOTESTAN(CIXI%)
               If LOTESTA < 1 Then LOTESTA = 1
            CICLOHE = !PieCiclo
               If CICLOHE < 1 Then CICLOHE = 1
            HORPREP = !HorsPrep
            HORFIJO = !HorsFijo
            MINVARI = !MinuStan
            HORHERR = !HorsMHer
            PIECICL = !PieCiclo
            HORLIMP = !HorsLimp
            DEMODIS = !PorDemor
            SOLAPAI% = !CoefSolp
            '
            CAOPS = !CantiOps
            CATOP% = !CategOps
            EQUIPO$ = !CodMaq0
            EQUIP1$ = !CodMaq1
            EQUIP2$ = !CodMaq2
            EQUIP3$ = !CodMaq3
            '
            OBSOPX$ = !OBSEOPER
            '
            PROTERI% = !ProTerOp
            PRECTER# = !PreUnid
            MONETER% = !MonePrec
            FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
            '
            X$ = REBLA$
            Call REPLA(X$, MKI$(CI%), 1, 2)
            Call REPLA(X$, MKI$(NOPY%), 3, 2)
            Call REPLA(X$, DOPE$, 5, 48)
            Call REPLA(X$, EQUIPO$, 53, 6)
            Call REPLA(X$, MKS$(HORPREP), 59, 4)
            Call REPLA(X$, MKS$(HORFIJO), 63, 4)
            Call REPLA(X$, MKS$(MINVARI), 67, 4)
            Call REPLA(X$, MKS$(DEMODIS), 75, 4)
            Call REPLA(X$, MKS$(HORLIMP), 79, 4)
            Call REPLA(X$, MKS$(CAOPS), 83, 4)
            Call REPLA(X$, Chr$(CATOP%), 87, 1)
            Call REPLA(X$, MKI$(SOLAPAI%), 89, 2)
            Call REPLA(X$, MKS$(LOTESTA), 91, 4)
            Call REPLA(X$, MKS$(HORHERR), 95, 4)
            Call REPLA(X$, MKS$(PIECICL), 99, 4)
            Call REPLA(X$, EQUIP1$, 103, 6)
            Call REPLA(X$, EQUIP2$, 109, 6)
            Call REPLA(X$, EQUIP3$, 115, 6)
            '
            Call REPLA(X$, MKI$(PROTERI%), 149, 2)
            Call REPLA(X$, MKD$(PRECTER#), 151, 8)
            Call REPLA(X$, MKI$(MONETER%), 159, 2)
            Call REPLA(X$, MKI$(FEPRETER%), 161, 2)
            
            If NOPY% > ULTOP% Then ULTOP% = NOPY%
            NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
            SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
            List3.AddItem SECO$
            RESECO$(List3.ListCount - 1) = X$  ' Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0)
            DESOPE$(List3.ListCount - 1) = OBSOPX$
            '
          End If
          .MoveNext
        Loop
      End With
      '
'      RNC$ = RECFILT$("COSOPTER", 1, MKI$(CI%))
'      URN& = ULTREG&("COSOPTER")
      '
'      For U& = 1 To Len(RNC$) Step 4
'        RENOTA& = CVS(Mid$(RNC$, U&, 4))
'        If RENOTA& > 0 Then
'          If RENOTA& <= URN& Then
'            TTZZ$ = REGLEIDO$("COSOPTER", RENOTA&)
'            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
'            For KKI% = 1 To List3.ListCount
'              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
'                Mid$(RESECO$(KKI% - 1), 145, 32) = TTZZ$
'              End If
'            Next KKI%
'          End If
'        End If
'      Next U&
      '
'      RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
'      URN& = ULTREG&("OBSEOPER")
'      '
'      For U& = 1 To Len(RNC$) Step 4
'        RENOTA& = CVS(Mid$(RNC$, U&, 4))
'        If RENOTA& > 0 Then
'          If RENOTA& <= URN& Then
'            TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
'            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
'            For KKI% = 1 To List3.ListCount
'              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
'                DESOPE$(KKI% - 1) = DESOPE$(KKI% - 1) + Mid$(TTZZ$, 5)
'              End If
'            Next KKI%
'          End If
'        End If
'      Next U&
      '
25    If List3.ListCount > 0 Then
        List3.ListIndex = 0
      End If
      '
      YACAR%(3) = 1
      MODIFIC%(3) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub MUEDATOPER()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   '
   RETEMODI3% = MODIFIC%(3)
   '
   CABAS = LOTESTAN(CI%)
   If CABAS < 1 Then CABAS = 1
   If CABAS > 3000000 Then CABAS = 1
   '
   REGI& = Val(Mid$(List3.TEXT, 59, 6))
   NOPY% = Val(Left$(List3.TEXT, 4))
   DOPE$ = Mid$(List3.TEXT, 10, 48)
   '
   VDEF$ = RESECO$(List3.ListIndex)
   '
   Label55.Caption = TRIM$(Str$(NOPY%))
   Text5.TEXT = DOPE$
   Text7.TEXT = CSTRING(MKS$(CABAS), 3, 7, 0, 2)
   '
   Text9(0).TEXT = AJUSTI$(Mid$(VDEF$, 53, 6), 6)
   Label43.Caption = ECOARCH$("PADMAQ", Text9(0).TEXT)
   Text9(1).TEXT = Mid$(VDEF$, 103, 6)
   Text9(2).TEXT = Mid$(VDEF$, 109, 6)
   Text9(3).TEXT = Mid$(VDEF$, 115, 6)
   '
   Text11.TEXT = Mid$(VDEF$, 129, 16)
   '
   TPREP = CVS(Mid$(VDEF$, 59, 4))
   TSTT$ = HORMINSEDIEM$(60 * 60 * TPREP)
   Text15(0).TEXT = Left$(TSTT$, 3)
   Text17(0).TEXT = Mid$(TSTT$, 5, 2)
   Text26(0).TEXT = Mid$(TSTT$, 8, 7)
   '
   TFIJO = CVS(Mid$(VDEF$, 63, 4))
   TSTT$ = HORMINSEDIEM$(60 * 60 * TFIJO)
   Text15(1).TEXT = Left$(TSTT$, 3)
   Text17(1).TEXT = Mid$(TSTT$, 5, 2)
   Text26(1).TEXT = Mid$(TSTT$, 8, 7)
   '
   TISTA = CABAS * CVS(Mid$(VDEF$, 67, 4))
   TSTT$ = HORMINSEDIEM$(60 * TISTA)
   Text15(2).TEXT = Left$(TSTT$, 3)
   Text17(2).TEXT = Mid$(TSTT$, 5, 2)
   Text26(2).TEXT = Mid$(TSTT$, 8, 7)
   '
   TLIMP = CVS(Mid$(VDEF$, 79, 4))
   TSTT$ = HORMINSEDIEM$(60 * 60 * TLIMP)
   Text15(3).TEXT = Left$(TSTT$, 3)
   Text17(3).TEXT = Mid$(TSTT$, 5, 2)
   Text26(3).TEXT = Mid$(TSTT$, 8, 7)
   '
   TMHER = CVS(Mid$(VDEF$, 95, 4))
   TSTT$ = HORMINSEDIEM$(60 * 60 * TMHER)
   Text15(4).TEXT = Left$(TSTT$, 3)
   Text17(4).TEXT = Mid$(TSTT$, 5, 2)
   Text26(4).TEXT = Mid$(TSTT$, 8, 7)
   '
   Text29.TEXT = CSTRING$(Mid$(VDEF$, 75, 4), 3, 4, 1, 2)
   Text30.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 83, 4), 3, 5, 2, 2))
   Text33.TEXT = TRIM$(Str$(Asc(Mid$(VDEF$, 87, 1))))
   Text34.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 99, 4), 3, 5, 0, 2))
   Label51.Caption = ECOARCH$("CATOPER", Chr$(Text33.TEXT))
   '
   Text2.TEXT = RTrim$(DESOPE$(List3.ListIndex))
   '
   Label60.Caption = ""
   Label62.Caption = ""
   RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
   URN& = ULTREG&("DATADOPS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         If CVI(Left$(TTXX$, 2)) = CI% Then
           FEX = CVI(Mid$(TTXX$, 25, 2))
           Label60.Caption = FECHATEX$(FEX)
           Label62.Caption = TRIM$(Mid$(TTXX$, 37, 16))
         End If
       End If
     End If
   End If
   '
   NPRO% = CVI(Mid$(RESECO$(List3.ListIndex), 149, 2))
   PREC# = CVD(Mid$(RESECO$(List3.ListIndex), 151, 8))
   MONEI% = CVI(Mid$(RESECO$(List3.ListIndex), 159, 2))
   If MONEI% < 1 Then MONEI% = 1
   FECOI = CVI(Mid$(RESECO$(List3.ListIndex), 161, 2))
   '
   Label48.Caption = ""
   If NPRO% > 0 Then
     PROVEX$ = AJUSTI$(ECOARCH$("ACREDOR", MKI$(NPRO%)), 48) + Str$(NPRO%)
     Label48.Caption = PROVEX$
   End If
   '
   Text37.TEXT = CSTRING$(MKD$(PREC#), 3, 11, 4, 2)
   '
   If MONEI% > 0 Then
     If MONEI% < 256 Then
       MONEX$ = AJUSTI$(ECOARCH$("TAMONED", Chr$(MONEI%)), 48) + Str$(MONEI%)
       Label46.Caption = MONEX$
     End If
   End If
   '
   Label63 = FECHATEX$(FECOI)
   '
   MODIFIC%(3) = RETEMODI3%
   '
End Sub
'
Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If MODIFIC%(0) > 0 Then GoSub 90: Call GRAGENITEM: CAMBIOEXCEL = 0
   If MODIFIC%(1) > 0 Then GoSub 90: Call GRAESPECI
   If MODIFIC%(2) > 0 Or (SSTab1.Tab = 2 And GRABATODO% = 1) Then Command9.SetFocus: Call GRAFORMULA
   If MODIFIC%(3) > 0 Then GoSub 90: Call GRAHORUTA
   CAMBIOEXCEL = 0
   '
   If GRABATODO% = 1 Then
     If SSTab1.Tab = 3 Then Call GRACOSTAN
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   Exit Sub
   '
90 On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
Return
   '
End Sub
'
Sub GRACOSTAN()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If DERACCE%("GRACOSTA", "Grabar Costos Standard por Articulo") < 2 Then
      Exit Sub
   End If
   '
   RETRA& = CI%
   If RETRA& <= ULTREG&("RECOSCAL") Then
      XX$ = REGLEIDO$("MASTER", RETRA& + 1)
      TIMA% = Asc(Mid$(XX$, 93, 1))
      UMED$ = AJUSTI$(Mid$(XX$, 109, 2), 2)
      FECO% = HOY(HOS$)
      YY$ = String$(128, 0)
      Call REPLA(YY$, MKI$(CI%), 1, 2)
      Call REPLA(YY$, MKI$(TIMA%), 3, 2)
      Call REPLA(YY$, UMED$, 5, 2)
      Call REPLA(YY$, MKI$(FECO%), 7, 2)
      '
      Call REPLA(YY$, MKD$(Val(Label90) + Val(Label82)), 9, 8)
      Call REPLA(YY$, MKD$(Val(Label88) + Val(Label83)), 17, 8)
      Call REPLA(YY$, MKD$(Val(Label86) + Val(Label84)), 25, 8)
      Call REPLA(YY$, MKD$(Val(Label85)), 33, 8)
      Call REPLA(YY$, MKD$(Val(Label75) + Val(Label72)), 41, 8)
      Call REPLA(YY$, MKD$(Val(Label79) + Val(Label74)), 49, 8)
      Call REPLA(YY$, MKD$(0), 57, 8)
      Call REPLA(YY$, MKD$(Val(Label80)), 65, 8)
      Call REPLA(YY$, MKD$(Val(Label70)), 73, 8)
      Call REPLA(YY$, MKD$(Val(Label65)), 81, 8)
      Call REPLA(YY$, MKD$(Val(Label96)), 89, 8)
      Call REPLA(YY$, MKD$(Val(Label103)), 97, 8)
      Call REPLA(YY$, MKD$(Val(Label112)), 105, 8)
      Call REPLA(YY$, MKD$(Val(Label116)), 113, 8)
      Call REPLA(YY$, MKS$(Val(Label117)), 121, 4)
      Call REPLA(YY$, MKS$(Val(Label119)), 125, 4)
      Call GRAREG("RECOSCAL", YY$, RETRA&)
   End If
   '
End Sub
'
Sub GRAGENITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") < 2 Then
      Exit Sub
   End If
   '
   If YACAR%(0) < 1 Then    ' SOLO CAMBIO LA DESCRIPCION O EL NRO DE PLANO
      Screen.MousePointer = 11
      MODESCRI% = 0
      DE$ = DESCRIT0$(CI%)
      If TRIM$(Text3) <> "" Then
        If TRIM$(Text3) <> DE$ Or TRIM$(Text45) <> NUPLANO$(CI%) Then
          DE$ = TRIM$(Text3)
          MODESCRI% = 1
          Master.FindFirst ("CODINT = " & CI%)
          Master.Edit
          Master!DESCRIPCION = DE$
          GoTo 30
        End If
      End If
   End If
   '
   If YACAR%(0) < 1 Then Exit Sub
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(0) = 0
       MODIFIC(0) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   MODESCRI% = 0
   DE$ = DESCRIT0$(CI%)
   Master.FindFirst ("CODINT = " & CI%)
   Master.Edit
   '
   If TRIM$(Text3) <> "" Then
     If TRIM$(Text3) <> DE$ Then
       DE$ = TRIM$(Text3)
       MODESCRI% = 1
       Master!DESCRIPCION = DE$
     End If
   End If
   '
   TTITPE$ = TRIM$(Text24.TEXT)
   PESU = Val(TTITPE$)
   If InStr(TTITPE$, "/") > 0 Then
      TTQP$ = TTITPE$
      TTQP$ = REPLACAR$(TTQP$, "/", ".")
      PESU = (-1) * Abs(Val(TTQP$))
   End If
   Master!PESUNI = PESU
   '
   TIMA% = Val(Text6.TEXT)
   Master!TIMATE = TIMA%
   '
   GRUIVA% = Val(Text35.TEXT)
   Master!GRUPIVA = GRUIVA
   '
   CRIT% = Val(Text8.TEXT)
   Master!CRITREP = CRIT%
   '
   MONEI% = Val(Text13.TEXT)
   Master!MONECOS = MONEI%
   '
   TCOLOR$ = TRIM$(Text22.TEXT)
   Master!COLOR = TCOLOR$
   '
   TTALLE$ = TRIM$(Text23.TEXT)
   Master!TALLE = TTALLE
   '
   CCOMPO% = Val(Text25.TEXT)
   Master!DESCOMPO = CCOMPO%
   '
   STOMINI = Val(Text18)
   Master!STOMIN = STOMINI
   '
   LOREPO = Val(Text19)
   Master!LOREPOS = LOREPO
   '
   STOMAX = Val(Text20)
   Master!STOMAX = STOMAX
   '
   UNID$ = TRIM$(Text4.TEXT)
   Master!UMEDIDA = UNID$
   '
   COU = Val(Text12.TEXT)
   Master!COSUNI = COU
   '
   FECOSTX = Text14.TEXT
   If FECOSTX = "" Then FECOSTX = Now
   On Error Resume Next
   Master!FECOSTO = CDate(FECOSTX)
   On Error GoTo 0
   '
   PORNA = Val(Text28.TEXT)
   Master!PORNACIO = PORNA
   '
   NPRO% = Val(Text10.TEXT)
   Master!PROVHABI = NPRO%
   NPROA% = CVI(Mid$(REMAS$, 123, 2))
   '
   PHABIT%(CI%) = (-1)
   If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
     PHABIT%(CI%) = NPRO%
   End If
   '
   GRUCO% = Val(Text16.TEXT)
   'Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   '
   If Abs(Val(Text28)) > 320 Then
       Call MENSERR(24, "% Gastos de Nacionalización No Válido.")
     Else
       FCPI% = Int(100 * Val(Text28.TEXT) + 0.5)
'       Call REPLA(REMAS$, MKI$(FCPI%), 127, 2)
   End If
   '
   OBSERVIT$ = RTrim$(Text40)
   Master!Observa = OBSERVIT
   '
30 MODIPLAN% = 0: PLANUE$ = ""
   On Error Resume Next
   PLANUE$ = TRIM$(Master!NPLANO)
   If PLANUE$ <> TRIM$(Left$(Text45, 24)) Then
     MODIPLAN% = 1
     UPDAPLANOS% = 1
   End If
   Master!NPLANO = TRIM$(Left$(Text45, 24))
   Master.Update
   UPDATMASTER% = 1
   '
   If NPRO% <> NPROA% Then  ' cambio de proveedor habitual
     ZZX$ = REGBLAN$("RECAMPH")
     Call REPLA(ZZX$, MKI$(CI%), 1, 2)
     HOII% = HOY(HOS$)
     Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
     Call REPLA(ZZX$, MKI$(NPROA%), 5, 2)
     Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
     Call REPLA(ZZX$, USERTER$, 9, 24)
     Call REPLA(ZZX$, "Costos Standard", 33, 32)
     Call REGAPP("RECAMPH", ZZX$)
     Call CIERRARCH("RECAMPH")
   End If
   '
   'If MODESCRI% > 0 Then
   '   CODARTIC$ = AJUSTI$(Text1.TEXT, 16)
   '   LOTOTE& = 64 * ULTREG&("ECOMAST")
    '  FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   '   Dim EL As String * 2048
   '   TTXX$ = String$(LOTOTE&, 0)
   '   '
   '   For U& = 1 To 1 + LOF(FIECO%) / 2048
   '      Get #FIECO%, U&, EL$
   '      On Error Resume Next
   '      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
   '      On Error GoTo 0
   '   Next U&
   '   TTXX$ = Left$(TTXX$, LOTOTE&)
   '   '
   '   For U& = 1 To Len(TTXX$) Step 64
   '     If Mid$(TTXX$, U&, 16) = CODARTIC$ Then
   '       Mid$(TTXX$, U& + 16, 46) = AJUSTI$(DE$, 46)
   '       GoTo 40
   '     End If
   '   Next U&
   '   '
40 '   JJ& = 0
   '   For U& = 1 To Len(TTXX$) Step 2048
   '     EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
   '     If Len(EL$) < 2048 Then
   '       EL$ = Left$(EL$ + String$(2048, 0), 2048)
   '     End If
   '     JJ& = JJ& + 1
   '     Put #FIECO%, JJ&, EL$
   '   Next U&
   '   Close #FIECO%
   '   Call FRESHECHO("ECOMAST")
   '   '
   'End If
   '
   'If CI% > 1 Then     ' fuerza actualizacion de descripcion
   '   DDD = DESCRIT$(CI% - 1)
   'End If
   Call CIERRARCH("*.*")
   '
   Call BAJALDISCO
   Screen.MousePointer = 1
   MODIFIC%(0) = 0
   MODESCRI% = 0
   MODIPLAN% = 0
   '
End Sub
'
Sub GRAESPECI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("MODISPEC", "Modificación de Especificaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       MODIFIC(1) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("ESPECIFI")
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 62)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CI%) + RETEX$(LL%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ESPECIFI", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("ESPECIFI", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("ESPECIFI")
   MODIFIC(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAFORMULA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(2) < 1 Then Exit Sub
   If MODIFIC%(2) < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Fórmula del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(2) = 0
       MODIFIC(2) = 0
       Call CARGENITEM
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   CAIT% = List1.ListCount
   '
   Dim EXPLOLI1$(256)
   CAIT1% = CAIT%
   For JJI% = 1 To 256
      EXPLOLI1$(JJI%) = ""
      If JJI% <= List1.ListCount Then
         EXPLOLI1$(JJI%) = List1.List(JJI% - 1)
      End If
   Next JJI%
   '
   For I& = 1 To CAIT%
     EXPLOLIN(I&) = List1.List(I& - 1)
     CICOMPO% = CODINT%(Left$(EXPLOLIN(I&), 16))
     Call EXPLOMUL(CICOMPO%)
     For JJI% = 1 To CONTAMUL%
       If CIK%(JJI%) = CI% Then
         Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".")
         YACAR%(2) = 0
         MODIFIC(2) = 0
         SSTab1.Tab = 1
         Call CARGENITEM
         GoTo 99
       End If
     Next JJI%
   Next I&
   '
   CAIT% = CAIT1%
   For JJI% = 0 To 256
      EXPLOLIN$(JJI%) = EXPLOLI1$(JJI%)
   Next JJI%
   '
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   CABAS = 1
   CAIT% = List1.ListCount
   '
   If CABAS > 0.01 Then
      Call EXPLOGR(CODD$, CABAS)
      MODIFIC%(2) = 0
   End If
   '
99 Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   Command18.Enabled = True
   Command19.Enabled = True
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAHORUTA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(3) < 1 Then MODIFIC%(3) = 0
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If DERACCE%("MODIRUTA", "Modificación Hojas de Ruta") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       MODIFIC%(3) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   Call ABRESECOP
   '
   With Secoper
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 66
     End If
     YAPRIMO% = 0
     ABUS$ = "CodIConj =" & CI% & " AND StatOper > 0"
     On Error Resume Next
62   If YAPRIMO% = 0 Then
         .FindFirst ABUS$
         YAPRIMO% = 1
       Else
621      .FindNext ABUS$
     End If
     If .NoMatch = False Then
       If IsNull(!AltOper) = True Or !AltOper < 1 Then
         .Edit
         !statoper = (-1)
         .Update
         GoTo 62
       End If
       GoTo 621
     End If
   End With
   '
66 With Secoper
   For LL% = 1 To List3.ListCount
    NOPYI% = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
    CANPRE = 0: CATPRE% = 0
     .FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
     If .NoMatch = False Then
       CANPRE = !NUOPER_PREP
       CATPRE% = !CATOPER_PREP
     End If
     '
     .AddNew
     !CODICONJ = CI%
     !CODXCONJ = CODEXT$(CI%)
     !NumeOper = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
     !DescOper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
     !statoper = 1
     !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
     !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
     !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
     !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
     !HorsPrep = CVS(Mid$(RESECO$(LL% - 1), 59, 4))
     !HorsFijo = CVS(Mid$(RESECO$(LL% - 1), 63, 4))
     !MinuStan = CVS(Mid$(RESECO$(LL% - 1), 67, 4))
     !HorsLimp = CVS(Mid$(RESECO$(LL% - 1), 79, 4))
     !HorsMHer = CVS(Mid$(RESECO$(LL% - 1), 95, 4))
     !PieCiclo = CVS(Mid$(RESECO$(LL% - 1), 99, 4))
     !PorDemor = CVS(Mid$(RESECO$(LL% - 1), 75, 4))
     If !MinuStan > 0.005 Then
       !Cadencia = 60 / !MinuStan
     End If
     !CantiOps = CVS(Mid$(RESECO$(LL% - 1), 83, 4))
     !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
     !CoefSolp = CVI(Mid$(RESECO$(LL% - 1), 89, 2))
     !LOTESTAN = CVS(Mid$(RESECO$(LL% - 1), 91, 4))
     !ProgCNum = TRIM$(Mid$(RESECO$(LL% - 1), 121, 8))
     !HoProFab = TRIM$(Mid$(RESECO$(LL% - 1), 129, 16))
     !ProTerOp = CVI(Mid$(RESECO$(LL% - 1), 149, 2))
     !PreUnid = CVD(Mid$(RESECO$(LL% - 1), 151, 8))
     !MonePrec = CVI(Mid$(RESECO$(LL% - 1), 159, 2))
     !FechPrec = CVDAT(CVI(Mid$(RESECO$(LL% - 1), 161, 2)))
     '
     !OBSEOPER = DESOPE(LL% - 1)
     !NUOPER_PREP = CANPRE
     !CATOPER_PREP = CATPRE%
     .Update
     '
   Next LL%
   '
63 .FindFirst "CODICONJ = " & CI%
   If .NoMatch = False Then
     If !statoper = (-1) Then
       .Delete
       GoTo 63
     End If
   End If
   '
64 .FindNext "CODICONJ = " & CI%
   If .NoMatch = False Then
     If !statoper = (-1) Then
       .Delete
       GoTo 63
     End If
     GoTo 64
   End If
   '
   .Close
   End With   ' secoper
   '
   MODIFIC(3) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRADATOPER(CI%, NOPE%, NHPF$)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         URN& = ULTREG&("DATADOPS")
         RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
         KK& = (-3)
         '
         TTYY$ = REGBLAN$("DATADOPS")
         Call REPLA(TTYY$, MKI$(CI%), 1, 2)
         Call REPLA(TTYY$, MKI$(NOPE%), 3, 2)
         Call REPLA(TTYY$, NHPF$, 5, 16)
         'Call REPLA(TTYY$, Text2.TEXT, 21, 4)
         FEXI% = HOY(HOS$)
         Call REPLA(TTYY$, MKI$(FEXI%), 25, 2)
         Call REPLA(TTYY$, MKI$(FEXI%), 27, 2)
         'Call REPLA(TTYY$, Text5.TEXT, 29, 8)
         USUA$ = USERNAME$ 'TRIM$(Text6.TEXT): If USUA$ = "" Then USUA$ = USERNAME$
         Call REPLA(TTYY$, USUA$, 37, 16)
         'Call REPLA(TTYY$, Text4(0).TEXT, 53, 48)
         '
15       KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
              RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Else
              URN& = URN& + 1
              RENOTA& = URN&
              If TRIM$(NHPF$) = "" Then
                If KK& > 1 Then
                  GoTo 20
                End If
              End If
         End If
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         NOPI% = CVI(Mid$(TTXX$, 3, 2))
         If NOPI% > 0 Then
            If NOPI% <> NOPE% Then
               GoTo 15
            End If
         End If
         Call GRAREG("DATADOPS", TTYY$, RENOTA&)
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
            If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
               Call GRAREG("DATADOPS", MKI$(CI%) + MKI$(0) + Space$(124), RENOTA&)
            End If
         Wend
         '
         Call CIERRARCH("DATADOPS")
      End If
   End If
   '
End Sub

Private Sub List4_DblClick()
   FCRP% = FOCOREP%
   FOCOREP% = 2
   CMONE$ = Mid$(List4.TEXT, 25, 2)
   If CMONE$ = "  " Or CMONE$ = "$ " Then
       Text21 = Mid$(List4.TEXT, 27, 10)
       Text14 = Left$(List4.TEXT, 8)
       MONEI% = Val(Text13.TEXT)
       If MONEI% <= 1 Then
         Text12 = ""
         FOCOREP% = 1
         Text12 = Mid$(List4.TEXT, 27, 10)
       End If
     Else
       KMONE% = 2
       For UV& = 1 To ULTREG&("TAMONED")
         XXV$ = REGLEIDO$("TAMONED", UV&)
         If Mid$(XXV$, 2, 2) = CMONE$ Then
           KMONE% = Asc(Left$(XXV$, 1))
           GoTo 19
         End If
       Next UV&
19     Text13 = TRIM$(Str$(KMONE%))
       Text12 = ""
       FOCOREP% = 1
       Text12 = Mid$(List4.TEXT, 27, 10)
       Text14 = Left$(List4.TEXT, 8)
   End If
   FOCOREP% = FCRP%
End Sub

Private Sub Listad_Click()
Call OPLIMAS04.Command2_Click
End Sub

Private Sub MaePro_Click()
Call Command45_Click
End Sub

Private Sub NumPro_Click()
Call Command5_Click
End Sub

Private Sub Pad_Click()
Call Command45_Click
End Sub

Private Sub Otro_Click()
Call ACCDIR("ANACOS")
End Sub

Private Sub Sal_Click()
Call Command3_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      If InStr(SSTab1.Caption, ". . .") > 0 Then
        SSTab1.Tab = PreviousTab
        Exit Sub
      End If
      '
      If MODOEDIT% = 1 Then
        If PreviousTab = 1 Then
          SSTab1.Tab = PreviousTab
        End If
        Call MENSAEDIT
        Exit Sub
      End If
      Call GRADATITEM
      If SSTab1.Tab = 3 Then
        If LLAMACOSTO$ <> "" Then
          If Text1.TEXT = LLAMACOSTO Then Text1.TEXT = ""
          Text1.TEXT = LLAMACOSTO$
          'LLAMACOSTO$ = ""
          Exit Sub
        End If
      End If
      '
      If SSTab1.Tab = 2 Then
        If LLAMACOSTO$ <> "" Then
          If PreviousTab = 3 Then
            List3.Clear
            YACAR%(3) = 0
          End If
        End If
      End If
      '
      Call CARDATITEM
      '
   End If
End Sub

Private Sub Susticod_Click()
   Call Command23_Click
End Sub

Private Sub StoLot_Click()
Call OPLIMAS04.Command7_Click
End Sub

Private Sub SustCod_Click()
Call Command23_Click
End Sub

Private Sub Text1_Change()
   '
   For KKII% = 0 To 15: YACAR%(KKII%) = 0: Next KKII%
   '
   Call BLANFORMU
   CIXII% = CODINT%(Text1.TEXT)

If CIXII% > 0 Then
   
      
      '
      GRCV% = GRUCOVEI%(CIXII%)
      If GRCV% < 1 Or GRCV% > 255 Then GRCV% = 0
      Text41 = FORMATNUM$(PORGAGE(1, GRCV%), "F5.1")
      Text36 = FORMATNUM$(PORGAGE(2, GRCV%), "F5.1")
      Text38 = FORMATNUM$(PORGAGE(3, GRCV%), "F5.1")
      Text42 = FORMATNUM$(PORGAGE(4, GRCV%), "F4.1")
      Text43 = FORMATNUM$(PORGAGE(5, GRCV%), "F4.1")
      Text39 = FORMATNUM$(PORGAGE(6, GRCV%), "F4.1")
      Text44 = FORMATNUM$(PORGAGE(7, GRCV%), "F4.1")
      '
      For KKII% = 0 To 7
        MODOCALCU%(KKII%) = MODOCA%(KKII%, GRCV%)
      Next KKII%
      '
      Call CARDATITEM
      
      'Aca valido para habilitar planilla Excel
      Call ValidaExcel
      
      
      '
      PRELISX# = PRELISTA#("LIPRE001", CIXII%) * TICAMONE(MONELISTA%(1))
      Label114.Caption = TRIM$(CSTRING$(MKD$(PRELISX#), 3, 12, DECICALCOS%, 2))
      Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Text44) / 100), 3, 12, DECICALCOS%, 2))
      Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
      PREVENETO# = XVALO(Label112)
      '
      Call Text41_Change
      Call Text36_Change
      Call Text38_Change
      Call Text42_Change
      Call Text43_Change
      Call Text39_Change
      DoEvents
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
End Sub

Private Sub Text10_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text10_DblClick()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text10.TEXT))))
   End If
End Sub


Private Sub Text10_LostFocus()
   PROVEHABI = Val(Text10.TEXT)
   If PROVEHABI < 1 Or PROVEHABI > 32767 Then
       Text10.TEXT = ""
     ElseIf ECOARCH$("ACREDOR", MKI$(Int(PROVEHABI))) = "" Then
       Text10.TEXT = ""
     Else
       Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(PROVEHABI)))
   End If
End Sub

Private Sub Text11_Change()
   If List3.ListIndex >= 0 Then
     Mid$(RESECO$(List3.ListIndex), 129, 16) = AJUSTI$(Text11.TEXT, 16)
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11.TEXT)
End Sub

Private Sub Text12_Change()
   If FOCOREP% <= 1 Then
     MONEI% = Val(Text13.TEXT)
     COSUDOL = Val(Text12)
     If MONEI% > 1 Then
         VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
       Else
         FEBAS% = FECHANUM(Text14)
         FEACT% = HOY(HOS$)
         VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
     End If
     Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text12_GotFocus()
   FOCOREP% = 1
End Sub

Private Sub Text12_LostFocus()
   Text12.TEXT = CSTRING$(MKS$(Val(Text12.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text13_Change()
   VALONU = Val(Text13.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text13.TEXT = "": Exit Sub
   End If
   Label35.Caption = ECOARCH$("TAMONED", Chr$(VALONU))
   '
   MONEI% = Val(Text13.TEXT)
   If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text13_DBLCLICK()
   Call SELECHO("TAMONED")
   VALONUE$ = VALACT1$("TAMONED")
   If VALONUE$ <> "" Then
      Text13.TEXT = VALONUE$
   End If
End Sub

Private Sub Text14_Change()
   MODIFIC%(0) = 1
   If FECHANUM(Text14) > 0 Then
      CIXI% = CODINT%(Text1)
      MONEI% = Val(Text13)
      COSUDOL = Val(Text12)
      If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIXI%) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
      End If
      Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
   End If
End Sub

Private Sub Text14_DblClick()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text14.TEXT = VALONUE$
   End If
End Sub

Private Sub Text15_Change(Index As Integer)
   If Val(Text15(Index).TEXT) > 999 Then Text15(Index).TEXT = "999"
   If Val(Text15(Index).TEXT) < 0 Then Text15(Index).TEXT = "  0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index).TEXT)
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text15(Index).TEXT)), 3, 3, 0, 2))
   If TTTT$ = "" Then TTTT$ = "0"
   While Len(TTTT$) < 3: TTTT$ = " " + TTTT$: Wend
   Text15(Index).TEXT = TTTT$
End Sub

Private Sub Text17_Change(Index As Integer)
   If Val(Text17(Index).TEXT) > 59 Then Text17(Index).TEXT = "59"
   If Val(Text17(Index).TEXT) < 0 Then Text17(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text17_GotFocus(Index As Integer)
   Text17(Index).SelStart = 0
   Text17(Index).SelLength = Len(Text17(Index).TEXT)
End Sub

Private Sub Text17_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text17(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text17(Index).TEXT = TTTT$
End Sub

Private Sub Text18_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text18_LostFocus()
   Text18.TEXT = CSTRING$(MKS$(Val(Text18.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
   If Val(Text20.TEXT) < Val(Text19.TEXT) Then
     Text20.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text2_Change()
   On Error Resume Next
   DESOPE$(List3.ListIndex) = Text2
   On Error GoTo 0
   MODIFIC%(3) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_LostFocus()
   Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text21_Change()
   If FOCOREP% = 2 Then
     MONEI% = Val(Text13.TEXT)
     On Error Resume Next
     COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
     On Error GoTo 0
     Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text21_GotFocus()
   FOCOREP% = 2
End Sub

Private Sub Text21_LostFocus()
   Text21.TEXT = CSTRING$(MKS$(Val(Text21.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text22_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text22_DblClick()
   Call SELECHO("TACOLOR")
   COLONUE$ = VALACT1$("TACOLOR")
   If COLONUE$ <> "" Then
      Text22.TEXT = COLONUE$
   End If
End Sub

Private Sub Text23_Change()
  CAMBIOEXCEL = 1
  MODIFIC%(0) = 1
End Sub

Private Sub Text23_DblClick()
   Call SELECHO("TATALLES")
   TALLENUE$ = VALACT1$("TATALLES")
   If TALLENUE$ <> "" Then
      Text23.TEXT = TALLENUE$
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_LostFocus()
   Text24.TEXT = CSTRING$(MKS$(Val(Text24.TEXT)), 3, 10, 5, 2)
End Sub

Private Sub Text25_Change()
   VALONU = Val(Text25.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text25.TEXT = "": Exit Sub
   End If
   Label39.Caption = ECOARCH$("TACOMPO", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_DblClick()
   Call SELECHO("TACOMPO")
   VALONUE$ = VALACT1$("TACOMPO")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Text26_Change(Index As Integer)
   If Val(Text26(Index).TEXT) > 59.9999 Then Text26(Index).TEXT = "59.9999"
   If Val(Text26(Index).TEXT) < 0 Then Text26(Index).TEXT = "00.0000"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text26_GotFocus(Index As Integer)
   Text26(Index).SelStart = 0
   Text26(Index).SelLength = Len(Text26(Index).TEXT)
End Sub

Private Sub Text26_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text26(Index).TEXT)), 3, 7, 4, 2))
   While Len(TTTT$) < 7: TTTT$ = "0" + TTTT$: Wend
   If TTTT$ = "0000000" Then TTTT$ = "00.0000"
   Text26(Index).TEXT = TTTT$
End Sub

Private Sub Text28_Change()
   MONEI% = Val(Text13.TEXT)
   If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_LostFocus()
  If Val(Text13) <= 1 Then Text28 = ""
  Text28 = AJUSTD$(FORMATNUM$(Val(Text28), "F7.2"), 8)
End Sub

Private Sub Text29_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 75, 4) = MKS(Val(Text29.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Call CALCOSOPE
End Sub

Private Sub Text29_GotFocus()
   Text29.SelStart = 0
   Text29.SelLength = Len(Text29.TEXT)
End Sub

Private Sub Text29_LostFocus()
   Text29.TEXT = CSTRING$(MKS$(Val(Text29.TEXT)), 3, 4, 1, 2)
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text30_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 83, 4) = MKS(Val(Text30.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Call CALCOSOPE
End Sub

Private Sub Text30_GotFocus()
   Text30.SelStart = 0
   Text30.SelLength = Len(Text30.TEXT)
End Sub

Private Sub Text30_LostFocus()
   Text30.TEXT = CSTRING$(MKS$(Val(Text30.TEXT)), 3, 5, 2, 2)
End Sub

Private Sub Text33_GotFocus()
   Text33.SelStart = 0
   Text33.SelLength = Len(Text33.TEXT)
End Sub

'
Private Sub Text34_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 99, 4) = MKS(Val(Text34.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Call CALCOSOPE
End Sub

Private Sub Text34_GotFocus()
   Text34.SelStart = 0
   Text34.SelLength = Len(Text34.TEXT)
End Sub

Private Sub Text34_LostFocus()
   Text34.TEXT = TRIM$(CSTRING$(MKS$(Val(Text34.TEXT)), 3, 5, 0, 2))
End Sub

Private Sub Text31_DblClick()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
   '
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      CODY$ = TRIM$(Text31.TEXT)
      CIY% = CODINT%(CODY$)
      If CIY% > 0 Then
         CODE$ = CODY$: CIE% = CIY%
         Text31.TEXT = CODY$
         CODE$ = CODY$: CIE% = CIY%
         Label40.Caption = DESCRIT0$(CIE%)
         Label41.Caption = UNIMED$(CIE%)
         Text32.SetFocus
        Else
         Call MENSERR(24, "Código no Válido")
         Text31.TEXT = CODE$
      End If
   End If
   '
End Sub

Private Sub Text32_GotFocus()
   Text32.SelStart = 0
   Text32.SelLength = Len(Text32.TEXT)
End Sub

Private Sub Text33_Change()
   VALONU = Val(Text33.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text33.TEXT = "": Exit Sub
   End If
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 87, 1) = Chr$(VALONU)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Label51.Caption = ECOARCH$("CATOPER", Chr$(VALONU))
   Call CALCOSOPE
End Sub

Private Sub Text33_DblClick()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Text35_Change()
   VALONU = Val(Text35.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text35.TEXT = "": Exit Sub
   End If
   Label57.Caption = ECOARCH$("GRUPIVA", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text35_DblClick()
   Call SELECHO("GRUPIVA")
   VALONUE$ = VALACT1$("GRUPIVA")
   If VALONUE$ <> "" Then
      Text35.TEXT = VALONUE$
   End If
End Sub

Private Sub Text36_Change()
   If MODOCALCU%(2) = 1 Then
        Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Text36) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label68 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text36) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub Text36_GotFocus()
   Text36.SelStart = 0
   Text36.SelLength = Len(Text36.TEXT)
End Sub

Private Sub Text36_LostFocus()
   If Val(Text36) > 999.9 Then Text36 = "999.9"
   Text36 = CSTRING$(MKS$(Val(Text36)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENFAB", Text36.TEXT)
End Sub

Private Sub Text37_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 151, 8) = MKD(Val(Text37))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = Val(Mid$(Label46, 49, 6))
   TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   If FECHANUM(Label63) < 1 Then
     Label63 = FECHATEX$(HOY(HOS$))
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_LostFocus()
   Text37.TEXT = CSTRING$(MKS$(Val(Text37.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text38_Change()
   If MODOCALCU%(3) = 1 Then
        Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Text38) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label77 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text38) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub Text38_GotFocus()
   Text38.SelStart = 0
   Text38.SelLength = Len(Text38.TEXT)
End Sub

Private Sub Text38_LostFocus()
   If Val(Text38) > 999.9 Then Text38 = "999.9"
   Text38 = CSTRING$(MKS$(Val(Text38)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENADM", Text38.TEXT)
End Sub

Private Sub Text39_Change()
   PREVEN# = 0
   If (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(6) = 1 Then
        Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text39) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label107 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text39) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label119 = ""
   If Val(Text39) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Text39)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Text39_GotFocus()
   Text39.SelStart = 0
   Text39.SelLength = Len(Text39.TEXT)
End Sub

Private Sub Text39_LostFocus()
   If Val(Text39) > 99.9 Then Text39 = "99.9"
   Text39 = CSTRING$(MKS$(Val(Text39)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORUTILI", Text39.TEXT)
End Sub

Private Sub Text4_Change()
   Label37.Caption = Text4.TEXT
   Label30.Caption = ECOARCH$("UMEDIDA", Text4.TEXT)
   MODIFIC%(0) = 1
End Sub

Private Sub Text4_DblClick()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Text40_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text41_Change()
   If MODOCALCU%(1) = 1 Then
        Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Text41) / 100), 3, 12, DECICALCOS%, 2))
     Else
        Label93 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text41) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub Text41_GotFocus()
   Text41.SelStart = 0
   Text41.SelLength = Len(Text41.TEXT)
End Sub

Private Sub Text41_LostFocus()
   If Val(Text41) > 999.9 Then Text41 = "999.9"
   Text41 = CSTRING$(MKS$(Val(Text41)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENMAT", Text41.TEXT)
End Sub

Private Sub Text42_Change()
   PREVEN# = 0
   If (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(4) = 1 Then
        Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text42) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label110 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text42) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text42_GotFocus()
   Text42.SelStart = 0
   Text42.SelLength = Len(Text42.TEXT)
End Sub

Private Sub Text42_LostFocus()
   If Val(Text42) > 99.9 Then Text42 = "99.9"
   Text42 = CSTRING$(MKS$(Val(Text42)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORCOMIS", Text42.TEXT)
End Sub

Private Sub Text43_Change()
   PREVEN# = 0
   If (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(5) = 1 Then
        Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text43) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label104 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Text43) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text43_GotFocus()
   Text43.SelStart = 0
   Text43.SelLength = Len(Text43.TEXT)
End Sub

Private Sub Text43_LostFocus()
   If Val(Text43) > 99.9 Then Text43 = "99.9"
   Text43 = CSTRING$(MKS$(Val(Text43)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORIMPUE", Text43.TEXT)
End Sub

Private Sub Text44_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Text44) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text44_GotFocus()
   Text44.SelStart = 0
   Text44.SelLength = Len(Text41.TEXT)
End Sub

Private Sub Text44_LostFocus()
   If Val(Text44) > 99.9 Then Text42 = "99.9"
   Text44 = CSTRING$(MKS$(Val(Text44)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORDESCU", Text44.TEXT)
End Sub

Private Sub Text45_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text5_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 5, 48) = AJUSTI$(Text5.TEXT, 48)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5.TEXT)
End Sub

Private Sub Text5_LostFocus()
   ILISTA& = List3.ListIndex
   For UI& = 1 To List3.ListCount
     If Val(Left$(List3.List(UI& - 1), 4)) = Val(Label55) Then
       TTXT$ = List3.List(UI& - 1)
       List3.RemoveItem (UI& - 1)
       Call REPLA(TTXT$, Text5.TEXT, 10, 48)
       List3.AddItem TTXT$, UI& - 1
     End If
   Next UI&
   On Error Resume Next
   List3.ListIndex = ILISTA&
   On Error GoTo 0
End Sub

Private Sub Text6_Change()
   VALONU = Val(Text6.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text6.TEXT = "": Exit Sub
   End If
   Label29.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_DblClick()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Text16_Change()
   VALONU = Val(Text16.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text16.TEXT = "": Exit Sub
   End If
   Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(Int(VALONU)))
   MODIFIC%(0) = 1
End Sub

Private Sub Text16_DblClick()
   Call SELECHO("GRUCOVEN")
   VALONUE$ = VALACT1$("GRUCOVEN")
   If VALONUE$ <> "" Then
      Text16.TEXT = VALONUE$
   End If
End Sub

Private Sub Text7_Change()
   '
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     CABAS = Val(Text7.TEXT)
     If CABAS < 1 Then
        CABAS = 1
        Text7.TEXT = "      1"
     End If
     '
     TISTA = CABAS * CVS(Mid$(VDEF$, 67, 4))
     TSTT$ = HORMINSEDIEM$(60 * TISTA)
     Text15(2).TEXT = Left$(TSTT$, 3)
     Text17(2).TEXT = Mid$(TSTT$, 5, 2)
     Text26(2).TEXT = Mid$(TSTT$, 8, 7)
   End If
   '
   Call CALCOSOPE
   MODIFIC%(3) = 1
   '
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub Text8_Change()
   VALONU = Val(Text8.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text8.TEXT = "": Exit Sub
   End If
   Label32.Caption = ECOARCH$("TACRIRE", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text8_DblClick()
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Text9_Change(Index As Integer)
   If List3.ListIndex >= 0 Then
     If TRIM$(Text9(Index).TEXT) = "" Or ECOARCH$("PADMAQ", Text9(Index).TEXT) <> "" Then
        If Index = 0 Then
            Label43.Caption = ECOARCH$("PADMAQ", Text9(Index).TEXT)
            Mid$(RESECO$(List3.ListIndex), 53, 6) = AJUSTI$(Text9(Index).TEXT, 6)
          Else
            Mid$(RESECO$(List3.ListIndex), 97 + 6 * Index, 6) = AJUSTI$(Text9(Index).TEXT, 6)
        End If
     End If
   End If
   MODIFIC(3) = 1
   Call CALCOSOPE
End Sub

Private Sub Text9_DblClick(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER
   Master.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Codint = " & CODII%
     Master.FindFirst "Codint = " & CODII%
     If Master.NoMatch = False Then
       CODII% = CODII% + 1
       GoTo 20
     End If
   End If
   '
   With Master
      .AddNew
      !CODIGO = CODINUE$
      !DESCRIPCION = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATMASTER% = 1
   End With
   '
99 Screen.MousePointer = 1

End Sub
'
Sub COPYRUTA(CODORI%, CODEST%)
   '
   Dim REGTX$(256)
   Dim RESECOI$(256), DESOPEI$(256)
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   Text1 = CODOR$
   YACAR%(3) = 0
   Call CARHORUTA
   '
   CARESE% = List3.ListCount
   For KU% = 1 To CARESE%
     RESECOI$(KU% - 1) = RESECO$(KU% - 1)
     DESOPEI$(KU% - 1) = DESOPE$(KU% - 1)
   Next KU%
   '
   Text1 = CODES$
   List3.Clear
   For KU% = 1 To CARESE%
     RESECO$(KU% - 1) = RESECOI$(KU% - 1)
     DESOPE$(KU% - 1) = DESOPEI$(KU% - 1)
     '
     NOPY% = CVI(Mid$(RESECO$(KU% - 1), 3, 2))
     NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
     DOPE$ = Mid$(RESECO$(KU% - 1), 5, 48)
     SECO$ = NOPEX$ + "     " + DOPE$
     List3.AddItem SECO$
   Next KU%
   '
   YACAR%(3) = 1
   MODIFIC%(3) = 1
   GRABATODO% = 1
   Call GRAHORUTA
   '
   Text1 = ""
   YACAR%(3) = 0
   Text1 = CODES$
   Call CARHORUTA
   '
Exit Sub
   '
   RFFX$ = RECFILT$("PLACONTR", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PLACONTR", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PLACONTR")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PLACONTR", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PLACONTR")
     End If
     VALANT$ = REGLEIDO$("PLACONTR", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PLACONTR", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PLACONTR")
     Call GRAREG("PLACONTR", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PLACONTR")
   '
   RFFX$ = RECFILT$("PUEPUNTO", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PUEPUNTO", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PUEPUNTO")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PUEPUNTO", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PUEPUNTO")
     End If
     VALANT$ = REGLEIDO$("PUEPUNTO", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PUEPUNTO", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PUEPUNTO")
     Call GRAREG("PUEPUNTO", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PUEPUNTO")
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub COPYSPEC(CODORI%, CODEST%)
   '
   Dim REGTX$(1024)
   Screen.MousePointer = 11
   RFFX$ = RECFILT$("ESPECIFI", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("ESPECIFI", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("ESPECIFI")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("ESPECIFI", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("ESPECIFI")
     End If
     VALANT$ = REGLEIDO$("ESPECIFI", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("ESPECIFI", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("ESPECIFI")
     Call GRAREG("ESPECIFI", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("ESPECIFI")
   Screen.MousePointer = 1
   '
End Sub
'
Sub REVIMASTER()                         ' revision del maestro
        '
Exit Sub
        Screen.MousePointer = 11
        HAYINCO% = 0
        '
        Call BLOQARCH("MASTER")
        Call BLOQARCH("INVERT")
        Call BLOQARCH("INVENUE")
        '
        Dim CODISEL$(32767), CISEL%(32767)
        '
        If Left$(UCase$(EMPREAC$), 5) = "BOREN" Then
           ' ejecuta conversion de codigos
           FIN& = NUMAS%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("MASTER", I& + 1)
              CODES$ = Mid$(X$, 77, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 77, 16)
                 Call GRAREG("MASTER", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("INVERT", I& + 1)
              CODES$ = Mid$(X$, 1, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 1, 16)
                 Call GRAREG("INVERT", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("INVERT")
           '
        End If
        '
        NM% = ULTREG&("MASTER") - 1
        If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
              X$ = REGLEIDO$("MASTER", 1&)
              Call REPLA(X$, MKI$(NM%), 1, 2)
              Call GRAREG("MASTER", X$, 1&)
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Indice Restaurado.")
              HAYINCO% = 1
          Else
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Control O.K.")
        End If
        '
        Screen.MousePointer = 11
        NVL& = ULTREG&("INVERT") - 1
        FIN& = NVL&
        For II& = 1 To NVL&
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("INVERT", II& + 1)
            COD$ = Left$(X$, 16): CII% = CVI(Mid$(X$, 17, 2))
            If Mid$(COD$, 16, 1) = Chr$(96) Then
               If CII% > 0 Then
                  If CII% <= NUMAS% Then
                     REGMA& = CII% + 1
                     Y$ = REGLEIDO$("MASTER", REGMA&)
                     If Mid$(Y$, 77, 15) = Left$(COD$, 15) Then
                        Call REPLA(Y$, Chr$(96), 92, 1)
                        Call GRAREG("MASTER", Y$, REGMA&)
                     End If
                  End If
               End If
            End If
        Next II&
        '
        Screen.MousePointer = 11
        FORMALTER.LINOSORT.Clear
        FORMALTER.LISORT.Clear
        FIN& = NM%
        RECUMAX& = 0
        For II& = 1 To NM%
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("MASTER", II& + 1)
            If CVI(Left$(X$, 2)) <> II& Then
                Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                Screen.MousePointer = 11
                cix% = II&
                Call REPLA(X$, MKI$(cix%), 1, 2)
                Call GRAREG("MASTER", X$, II& + 1)
            End If
            COD$ = AJUSTI$(TRIM$(Mid$(X$, 77, 15)), 15) + Mid$(X$, 92, 1)
            If Left$(COD$, 4) = "zzz-" Then
              If Val(Mid$(COD$, 5, 6)) > RECUMAX& Then
                RECUMAX& = Val(Mid$(COD$, 5, 6))
              End If
            End If
            '
            ZZ$ = "A" + COD$ + CSTRING$(MKS$(II&), 3, 7, 0, 2)
            FORMALTER.LISORT.AddItem ZZ$
        Next II&
        '
        FIN& = FORMALTER.LISORT.ListCount
        For II& = 1 To FORMALTER.LISORT.ListCount
           Call TRACE(20, II&, FIN&)
           ZZ$ = FORMALTER.LISORT.List(II& - 1)
           Call AGRELISCOD(ZZ$)
        Next II&
        '
        NVL& = ULTREG&("INVERT") - 1
        If NVL& < 1 Then
             Call REGAPP("INVERT", String$(16, 0) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = FORMALTER.LINOSORT.ListCount
        X$ = REGBLAN$("INVERT")
        Call REPLA(X$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", X$, 1)
        '
        j& = 1
        FIN& = NV% - 1
        For II& = 0 To FORMALTER.LINOSORT.ListCount - 1
            Call TRACE(20, II&, FIN&)
            FORMALTER.LINOSORT.ListIndex = II&
            CODISEL$(II&) = Mid$(FORMALTER.LINOSORT.TEXT, 2, 16)
            CISEL%(II&) = Val(Mid$(FORMALTER.LINOSORT.TEXT, 18))
            X$ = REGBLAN$("INVERT")
            Call REPLA(X$, CODISEL$(II&), 1, 16)
            Call REPLA(X$, MKI$(CISEL%(II&)), 17, 2)
            Call GRAREG("INVERT", X$, II& + 2)
        Next II&
        Call SETULTREG("INVERT", 1& + NV%)
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        CODA$ = ""
        FIN& = NV%
        For II& = 1 To NV%
           Call TRACE(20, II&, FIN&)
           IK$ = REGLEIDO$("INVERT", II& + 1)
           CODL$ = Left$(IK$, 15)
           If TRIM$(CODL$) = "" Or CODL$ = CODA$ Then
                'If TRIM$(CODL$) = "" Then GoTo 19
                REPLAI% = REPLAI% + 1
                'CODNUE$ = TRIM$(CODL$) + "-" + Chr$(64 + REPLAI%)
                CODNUE$ = "zzz-" + TRIM$(Str$(RECUMAX& + 1))
                RECUMAX& = RECUMAX& + 1
                If Left$(CODNUE$, 1) = "-" Then CODNUE$ = "0" + CODNUE$
                RGM& = 1 + CVI(Mid$(IK$, 17, 2))
                M$ = REGLEIDO$("MASTER", RGM&)
                Call REPLA(M$, CODNUE$, 77, 16)
                Call GRAREG("MASTER", M$, RGM&)
                Call REPLA(IK$, CODNUE$, 1, 15)
                Call REPLA(IK$, Chr$(96), 16, 1)
                Call GRAREG("INVERT", IK$, II& + 1)
                Call COMUNI("Codigo '" + TRIM$(CODL$) + "' Duplicado o no Válido.\  \Reemplazado por '" + TRIM$(CODNUE$) + "'.")
                Screen.MousePointer = 11
                HAYINCO% = 1
              Else
                REPLAI% = 0
           End If
           '
           If REPLAI% = 0 Then CODA$ = CODL$
19      Next II&
        '
        Call CIERRARCH("*.*")
        Call BLANARCH("INVENUE")
        '
17000   'Call BLIMESS("Actualizando Archivos de Validación ...")
        '
        'Call BLOQARCH("ECOMAST")
        FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 64)
        Dim ECM As String * 64
        '
        'Call BLOQARCH("CIACTI")
        FICIA% = FILEOPEN%(LUDAT$ + "CIACTI.DAT", 0, 2)
        Dim CCT As String * 2
        '
        NV% = ULTREG&("INVERT") - 1
        CCT$ = MKI$(NV%)
        Put #FICIA%, 1, CCT$
        '
        K% = 0
        FIN& = NV%
        For XL& = 1 To NV%
             Call TRACE(20, XL&, FIN&)
             IK$ = REGLEIDO$("INVERT", XL& + 1)
             If Mid$(IK$, 16, 1) <> Chr$(96) Then
               If TRIM$(Left$(IK$, 16)) <> "" Then
                 '
                 CI% = CVI(Mid$(IK$, 17, 2))
                 DE$ = DESCRIT$(CI%)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, Left$(IK$, 16), 1, 16)
                 Call REPLA(EC$, DE$, 17, 46)
                 '
                 K% = K% + 1: ECM$ = EC$: Put #FIECO%, K%, ECM$
                              CCT$ = MKI$(CI%): Put #FICIA%, K% + 1, CCT$
                 '
               End If
             End If
        Next XL&
        '
        CCT$ = MKI$(K%): Put #FICIA%, 1, CCT$
        '
        For K1% = K% + 1 To LOF(FIECO%) / 64
           ECM$ = String$(64, 0): Put #FIECO%, K1%, ECM$
           CCT$ = String$(2, 0): Put #FICIA%, K1% + 1, CCT$
        Next K1%
        '
        Close #FIECO%: FIECO% = 0 'Call CIERRARCH("ECOMAST")
        Close #FICIA%: FICIA% = 0 'Call CIERRARCH("CIACTI")
        FIECO% = 0: FICIA% = 0
        '
        Call GRACONTROL("", "AMASTO", "NORMAL")
        '
        Call CIERRARCH("*.*")
        Call LIBARCH("MASTER")
        '
        Call FRESHECHO("MASTER")
        Call FRESHECHO("ECOMAST")
        '
        'Call BLIMESS("-")
        '
        Screen.MousePointer = 1
        If HAYINCO% = 0 Then
            Call COMUNI("No se han Detectado Inconsistencias\en la Revisión del Maestro de Items.")
        End If
        Exit Sub
        '
        '
12500   If ULTREG&("INVERT") < 1 Then
             Call REGAPP("INVERT", String$(15, 0) + Chr$(1) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = ULTREG&("INVERT") - 1
        If CVI(Mid$(REGLEIDO$("INVERT", 1&), 17, 2)) <> NV% Then
            Call COMUNI("Posible Inconsistencia (1) !!!\Ejecute Opción de Control.")
            Call FINAL("")    ' GoTo 2000
        End If
        '
        ABRE$ = AJUSTI$(COD$, 15)
        LI& = 1: LS& = NV% + 2
        ENCU% = 0
        '
12510   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 12580
        L& = LI& + E&: If L& < 2 Or L& > NV% + 1 Then GoTo 12580
        IK$ = REGLEIDO$("INVERT", L&): I1$ = Left$(IK$, 15)
                                  If I1$ < ABRE$ Then LI& = L&: GoTo 12510
                                  If I1$ > ABRE$ Then LS& = L&: GoTo 12510
        ENCU% = 1: CI% = CVI(Mid$(IK$, 17, 2)): RGI& = L&
                If Mid$(IK$, 16, 1) = Chr$(96) Then
                    ENCU% = (-1)
                End If
        '
12580   Return
        '
        '
End Sub

Sub AGRELISCOD(RAGRE$)
   '
   ULIS& = FORMALTER.LINOSORT.ListCount
   If ULIS& < 1 Or Left$(RAGRE$, 17) > FORMALTER.LINOSORT.List(ULIS& - 1) Then
       FORMALTER.LINOSORT.AddItem RAGRE$
       GoTo 10
     Else
       For j& = ULIS& - 1 To 0 Step -1
         If Left$(FORMALTER.LINOSORT.List(j&), 17) <= Left$(RAGRE$, 17) Then
           FORMALTER.LINOSORT.AddItem RAGRE$, j& + 1
           GoTo 10
         End If
       Next j&
       FORMALTER.LINOSORT.AddItem RAGRE$, 0
10 End If
   '
End Sub

Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta Definir Depósitos")
        'Call CONECRUN("FLSETUP/ANACOSTO", "")
        OpAnCoAcc.Show 1
        GoTo 10
    End If
    '
    If ULTREG&("UMEDIDA") < 2 Then
        Call BLANARCH("UMEDIDA")
        Call REGAPP("UMEDIDA", "U.Unidades")
        Call REGAPP("UMEDIDA", "KgKilogramos")
        Call REGAPP("UMEDIDA", "grGramos")
        Call REGAPP("UMEDIDA", "gKGramos - Kilos")
        Call REGAPP("UMEDIDA", "M.Metros")
        Call REGAPP("UMEDIDA", "M2M.Cuadrados")
        Call REGAPP("UMEDIDA", "M3M.Cúbicos")
        Call REGAPP("UMEDIDA", "mmMilímetros")
        Call REGAPP("UMEDIDA", "cmCentímetros")
        Call REGAPP("UMEDIDA", "LtLitros")
        Call CIERRARCH("UMEDIDA")
    End If
    XX$ = REGLEIDO$("UMEDIDA", 1)
    '
    If ULTREG&("UNIMED") < 2 Then
        Call BLANARCH("UNIMED")
        Call REGAPP("UNIMED", "U.  Unidades" + String$(256, 0))
        Call REGAPP("UNIMED", "Kg  Kilogramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gr  Gramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gK  Gramos - Kilos" + String$(256, 0))
        Call REGAPP("UNIMED", "M.  Metros" + String$(256, 0))
        Call REGAPP("UNIMED", "M2  M.Cuadrados" + String$(256, 0))
        Call REGAPP("UNIMED", "M3  M.Cúbicos" + String$(256, 0))
        Call REGAPP("UNIMED", "mm  Milímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "cm  Centímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "Lt  Litros" + String$(256, 0))
        Call CIERRARCH("UNIMED")
    End If
    XX$ = REGLEIDO$("UNIMED", 1)
    '
    If ULTREG&("TATIMAT") < 1 Then
        Call BLANARCH("TATIMAT")
        Call REGAPP("TATIMAT", Chr$(1) + "Materia Prima")
        Call REGAPP("TATIMAT", Chr$(2) + "Pieza Comercial")
        Call REGAPP("TATIMAT", Chr$(3) + "Pieza Elaborada")
        Call REGAPP("TATIMAT", Chr$(4) + "Semielaborado")
        Call REGAPP("TATIMAT", Chr$(5) + "Prod.Terminado")
        Call REGAPP("TATIMAT", Chr$(8) + "Servicios")
        Call REGAPP("TATIMAT", Chr$(9) + "Improductivo")
        Call CIERRARCH("TATIMAT")
    End If
    XX$ = REGLEIDO$("TATIMAT", 1)
    '
    If ULTREG&("TABTIMA") < 1 Then
       DPPRE% = Asc(Left$(REGLEIDO$("ECODEP", 1), 1))
       For U& = 1 To ULTREG&("TATIMAT")
          X$ = REGLEIDO$("TATIMAT", U&)
          Y$ = REGBLAN$("TABTIMA")
          Call REPLA(Y$, X$, 1, 16)
          Call REPLA(Y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", Y$)
       Next U&
       Call CIERRARCH("TABTIMA")
    End If
    XX$ = REGLEIDO$("TABTIMA", 1)
    '
    If ULTREG&("TACRIRE") < 4 Then
        Call BLANARCH("TACRIRE")
        Call REGAPP("TACRIRE", Chr$(1) + "Por Pedido")
        Call REGAPP("TACRIRE", Chr$(2) + "Por Mínimos")
        Call REGAPP("TACRIRE", Chr$(3) + "Cons.Estadístico")
        Call REGAPP("TACRIRE", Chr$(4) + "Programación")
        Call CIERRARCH("TACRIRE")
    End If
    XX$ = REGLEIDO$("TACRIRE", 1)
    '
    If ULTREG&("TAMONED") < 2 Then
        Call BLANARCH("TAMONED")
        Call REGAPP("TAMONED", Chr$(1) + AJUSTI$("Pesos", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(2) + AJUSTI$("Dolares U$S", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(3) + AJUSTI$("Coef.Est.Ref.", 27) + MKS$(1))
        Call CIERRARCH("TAMONED")
    End If
    XX$ = REGLEIDO$("TAMONED", 1)
    '
    If ULTREG&("TACOMPO") < 2 Then
        Call BLANARCH("TACOMPO")
        Call REGAPP("TACOMPO", Chr$(0) + "Manual")
        Call REGAPP("TACOMPO", Chr$(1) + "Automática")
        Call CIERRARCH("TACOMPO")
    End If
    XX$ = REGLEIDO$("TACOMPO", 1)
    '
    If ULTREG&("GRUPIVA") < 1 Then
        Call SETULTREG("GRUPIVA", 0)
        Call REGAPP("GRUPIVA", Chr$(0) + "IVA 21 %   " + MKS$(21))
        Call REGAPP("GRUPIVA", Chr$(1) + "IVA 10,5 % " + MKS$(10.5))
        Call REGAPP("GRUPIVA", Chr$(2) + "IVA 27 %   " + MKS$(27))
        Call CIERRARCH("GRUPIVA")
    End If
    XX$ = REGLEIDO$("GRUPIVA", 1)
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub COPYSTRUC()
   '
10 CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     GoTo 99
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   Call LEEEXPLO(CODOR$, 1)
   If CAIT% < 1 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("Copia Imposible.\Código Origen " + TRIM$(CODOR$) + " sin Estructura.\\Repetir\Abandonar")
     If OPX% = 1 Then GoTo 10
     GoTo 99
   End If
   '
   Call LEEEXPLO(CODES$, 1)
   If CAIT% > 0 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("ATENCION: Código Destino " + TRIM$(CODES$) + " con Estructura.\\Cancelar\Sobre-escribir")
     If OPX% <> 2 Then GoTo 10
   End If
   '
   Call COPYFORMU(CODORI%, CODEST%)
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   '
99 End Sub

Sub COPYRUTAS()
   '
Dim REG&(256), REGTX$(256)
   '
10 Screen.MousePointer = 1
   CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     Call CIERRARCH("*.*")
     Exit Sub
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   Call ABRESECOP
   SQLX$ = "SELECT * FROM Secoper"
   SQLX$ = SQLX$ & " WHERE CodIConj = " & CODORI%
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   SecopTemp.MoveFirst
     If Err Then
       Screen.MousePointer = 1
       Call COMUNI("Copia Imposible.\Código Origen sin Operaciones.")
       GoTo 10
     End If
     '
   SQLX$ = "SELECT * FROM Secoper"
   SQLX$ = SQLX$ & " WHERE CodIConj = " & CODEST%
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   SecopTemp.MoveFirst
     If Err Then GoTo 20
   Screen.MousePointer = 1
   OPX% = COMALTER%("ATENCION: Código Destino con Operaciones !!!\\Cancelar\Sobre-escribir")
   If OPX% <> 2 Then GoTo 10
   '
20 Call COPYRUTA(CODORI%, CODEST%)
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'
Sub LISTRUTAS()
   '
   XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Dim PUNCOD$(32767)
   '
   CI1% = CVI(Left$(X$, 2))
   COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2))
   COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
   '
   SQLX$ = "SELECT * FROM Master "
   SQLX$ = SQLX$ + "WHERE CODIGO>='" & COD1$ & "' "
   SQLX$ = SQLX$ + "AND CODIGO<='" & COD2$ & "' "
   SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
   Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   MasterTemp.MoveLast
   If Err Then Exit Sub
   '
   CAREC% = MasterTemp.RecordCount
   If CAREC% < 1 Then Exit Sub
   '
   FIN& = CAREC%
   MasterTemp.MoveFirst: U& = 1
   '
   Call ABREMASTER
   '
10 Do Until MasterTemp.EOF = True
     Call TRACE(20, U&, FIN&)
     CI0% = MasterTemp!CODINT
     SQLX$ = "SELECT * FROM Secoper"
     SQLX$ = SQLX$ & " WHERE CodIConj = " & CI0%
     Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With SecopTemp
       .MoveFirst
       If Err Then GoTo 25
       '
       LOSTAND = LOTESTAN(CI0%): If LOSTAND < 1 Then LOSTAND = 1
       DESCRI$ = TRIM$("Lt.Std." + TRIM$(Str$(LOSTAND)) + " " + TRIM$(UNIMED$(CI0%)))
       DESCRI$ = " (" + DESCRI$ + ")"
       LOMADE% = 60 - Len(DESCRI$)
       If LOMADE% > 0 Then
          DESCRI$ = Left$(TRIM$(DESCRIT0$(CI0%)), LOMADE%) + DESCRI$
       End If
       '
       Y$ = REGBLAN$("SECOPSEL")
       Call REPLA(Y$, MKI$(CI0%), 1, 2)
       Call REPLA(Y$, DESCRI$, 3, 60)
       V& = V& + 1
       Call GRAREG("SECOPSEL", Y$, V&)
       '
       Call REPLA(Y$, Space$(60), 3, 60)
       V& = V& + 1
       Call GRAREG("SECOPSEL", Y$, V&)
       '
       Do Until .EOF = True
         If IsNull(!AltOper) = True Or !AltOper < 1 Then
           NOPY% = !NumeOper
           DOPE$ = !DescOper
           LOSTAND = LOTESTAN(CIXI%)
               If LOSTAND < 1 Then LOSTAND = 1
           CICLOHE = !PieCiclo
               If CICLOHE < 1 Then CICLOHE = 1
           HORPREP = !HorsPrep
           HORFIJO = !HorsFijo
           MINVARI = !MinuStan
           HORHERR = !HorsMHer
           PIECICL = !PieCiclo
           HORLIMP = !HorsLimp
           DEMODIS = !PorDemor
           SOLAPAI% = !CoefSolp
           '
           CAOPS = !CantiOps
           CATOP% = !CategOps
           EQUIPO$ = !CodMaq0
           EQUIP1$ = !CodMaq1
           EQUIP2$ = !CodMaq2
           EQUIP3$ = !CodMaq3
           '
           PROTERI% = !ProTerOp
           PRECTER# = !PreUnid
           MONETER% = !MonePrec
           FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
           '
           NOPES$ = TRIM$(FORMATNUM$(NOPY%, "I4"))
             While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
           DOPES$ = DOPE$
           TVAR = MINVARI * LOSTAND
           '
           Call REPLA(Y$, NOPES$, 3, 6)
           Call REPLA(Y$, DOPES$, 9, 54)
           Call REPLA(Y$, MKS$(HORPREP), 63, 4)
           Call REPLA(Y$, MKS$(HORFIJO), 67, 4)
           Call REPLA(Y$, MKS$(TVAR), 71, 4)
           Call REPLA(Y$, MKS$(DEMODIS), 75, 4)
           Call REPLA(Y$, MKS$(HORLIMP), 79, 4)
           Call REPLA(Y$, MKS$(CAOPS), 83, 4)
           Call REPLA(Y$, MKS$(LOSTAN), 91, 4)
           Call REPLA(Y$, EQUIP1$, 103, 6)
           Call REPLA(Y$, EQUIP2$, 109, 6)
           Call REPLA(Y$, EQUIP3$, 115, 6)
           '
           V& = V& + 1
           Call GRAREG("SECOPSEL", Y$, V&)
         End If
       .MoveNext
       Loop
       '
     End With
     '
     Call REPLA(Y$, String$(126, 0), 3, 126)
     V& = V& + 1
     Call GRAREG("SECOPSEL", Y$, V&)
     '
     Y$ = REGBLAN$("SECOPSEL")
     V& = V& + 1
     Call GRAREG("SECOPSEL", Y$, V&)
     '
25 MasterTemp.MoveNext
   U& = U& + 1
   Loop
   '
   Call SETULTREG("SECOPSEL", V&)
   Call CIERRARCH("SECOPSEL")
   Screen.MousePointer = 1
   Call FINAL("*LISECOP")
   '
End Sub

Sub REDESPIE()
    '
    CAERR% = 0
    FIN& = NUMAS%
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CIBAS% = I&
      '
      Call AREXPLO(CIBAS%, 1, 3)
      For j% = 1 To CAIT%
        For K% = j% + 1 To CAIT%
          If CIJ%(j%) = CIJ%(K%) Then
            COD1$ = TRIM$(CODEXT$(CIBAS%))
            COD2$ = TRIM$(CODEXT$(CIJ%(j%)))
            If CODINT%(COD1$) = CIBAS% Then ' SUPRIME ITEMS DE BAJA
              CAERR% = CAERR% + 1
              MENSA$ = "Repetición del Código " + COD2$ + "\en Fórmula de Codigo " + COD1$
              If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
            End If
          End If
        Next K%
      Next j%
      '
      Call EXPLOMUL(CIBAS%)
      For j% = 2 To CONTAMUL%
        If CIK%(j%) = CIBAS% Then
          CAERR% = CAERR% + 1
          MENSA$ = "Recursividad en Código " + TRIM$(CODEXT$(CIBAS%)) + "\" + DESCRIT0$(CIBAS%)
          If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
        End If
      Next j%
      '
89  Next I&
    '
    If CAERR% < 1 Then
        Call COMUNI("Control O.K. - No se Detectaron Repeticiones")
      Else
        Call MENSERR(20, "Se Encontraron" + Str$(CAERR%) + " Repeticiones en Despieces")
    End If
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub


Sub LIESTRU()      ' listado estructuras y enlaces por codigo
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    CI1% = CVI(Left$(X$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    FIECO% = FILENBR%("LIESTRU")
    Dim ECM As String * 16
    '
    Call HEADERS("LIESTRU", "")
    Call HEADERS("LIESTRU", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LIESTRU", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '
    Screen.MousePointer = 11
    FIN& = NUMAS%
    j& = 0
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CI% = I&
      If STATUIT%(CI%) > 0 Then
        ACTIX% = 0
        CODYY$ = CODEXT$(CI%)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            PORMILLAR% = 0
            Call LEEEXPLO(CODYY$, 1)
            If CAIT% > 0 Then
              For K% = 1 To CAIT%
                CI1% = CIJ%(K%)
                CAN = USOI(K%)
                PD% = PDI%(K%)
                If CI1% > 0 Then
                  If CI1% <= NUMAS% Then
                    If Abs(CAN) > 0.00005 Then
                      UM$ = UNIMED$(CI1%)
                      Z$ = REGBLAN$("LIESTRU")
                      Call REPLA(Z$, MKI$(CI%), 1, 2)
                      Call REPLA(Z$, MKI$(CI1%), 3, 2)
                      Call REPLA(Z$, UM$, 5, 2)
                      Call REPLA(Z$, MKS$(CAN), 7, 4)
                      Call REPLA(Z$, Chr$(PD%), 11, 1)
                      '
                      j& = j& + 1
                      Call GRAREG("LIESTRU", Z$, j&)
                      ACTIX% = 1
                      '
                    End If
                  End If
                End If
              Next K%
            End If
          End If
        End If
        '
        If ACTIX% > 0 Then
          Z$ = REGBLAN$("LIESTRU")
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          j& = j& + 1
          Call GRAREG("LIESTRU", Z$, j&)
        End If
        '
      End If
      '
    Next I&
    '
    Call SETULTREG("LIESTRU", j&)
    Call CIERRARCH("LIESTRU")
    Call FILESORT("LIESTRU", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    Call FINAL("*LIESTRU")
    '
End Sub
'
Sub LIMPLOSI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
60 Screen.MousePointer = 11
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Implosión " + DESCRIT$(CI%)
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left + (Width - ScaleWidth) / 2
   IMPLOSI.Show
   '
   Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas WHERE FORMULAS!CODIELEM=" & CI%, dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   FORMULASTEMP.MoveLast
   If Err Then GoTo 99
   CAREC% = FORMULASTEMP.RecordCount
   If CAREC% < 1 Then GoTo 99
   '
   FORMULASTEMP.MoveFirst
   ACTIREC% = 1
61 If FORMULASTEMP.NoMatch = False Then
      CICONJU% = FORMULASTEMP!CODICONJ
      USOX = FORMULASTEMP!USOBRUTO
      If CICONJU% > 0 Then
        If CICONJU% <= NUMAS% Then
          If PORMILLAR% < 1 Then
            If HABIPORMI% > 0 Then
              If USOX > 0 Then
                If USOX < 0.005 Then
                  PORMILLAR% = 1
                  GoTo 60
                End If
              End If
            End If
          End If
          If PORMILLAR% = 1 Then USOX = 1000 * USOX
          Z$ = Space$(64)
          Call REPLA(Z$, CODEXT$(CICONJU%), 1, 16)
          Call REPLA(Z$, DESCRIT0$(CICONJU%), 17, 32)
          Call REPLA(Z$, UNIMED$(CICONJU%), 51, 2)
          Call REPLA(Z$, FORMATNUM$(USOX, "F10.4"), 55, 10)
          IMPLOSI.List1.AddItem Z$
          On Error Resume Next
          IMPLOSI.List1.TopIndex = IMPLOSI.List1.ListCount - 10
          On Error GoTo 0
          DoEvents
        End If
      End If
      '
      If ACTIREC% < CAREC% Then
        ACTIREC% = ACTIREC% + 1
        FORMULASTEMP.MoveNext
        GoTo 61
      End If
   End If
   '
   NOCHECK% = 1
   Check1.Value = 0
   If PORMILLAR% = 1 Then NOCHECK% = 1: Check1.Value = 1
   NOCHECK% = 0
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub LIMULTI()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim TIESTRU%(32767), ESPARSTRU%(32767)
    '
    NMS% = NUMAS%
    FIN& = NUMAS%
    j& = 0
    '
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CI% = I&
      If STATUIT%(CI%) > 0 Then
        CODYY$ = CODEXT$(CI%)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            PORMILLAR% = 0
            Call LEEEXPLO(CODYY$, 1)
            If CAIT% > 0 Then
              TIESTRU%(CI%) = 1
              For K% = 1 To CAIT%
                CI1% = CIJ%(K%)
                If CI1% > 0 Then
                  If CI1% <= NMS% Then
                    ESPARSTRU%(CI1%) = 1
                  End If
                End If
              Next K%
            End If
          End If
        End If
      End If
    Next I&
    '
    SQLX$ = "SELECT * FROM Master "
    SQLX$ = SQLX$ + "WHERE CODIGO>='" & COD1$ & "' "
    SQLX$ = SQLX$ + "AND CODIGO<='" & COD2$ & "' "
    SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
    Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    MasterTemp.MoveLast
    If Err Then Exit Sub
    '
    CAREC% = MasterTemp.RecordCount
    If CAREC% < 1 Then Exit Sub
    '
    FIN& = CAREC%
    MasterTemp.MoveFirst: U& = 1
    '
10  Do Until MasterTemp.EOF = True
      Call TRACE(20, U&, FIN&)
      CI% = MasterTemp!CODINT
      If (TIESTRU%(CI%) = 1 And ESPARSTRU%(CI%) = 0) Or CI1% = CI2% Then
        Call EXPLOMUL(CI%)
        If CONTAMUL% > 1 Then
          '
          If REGQ& > 0 Then
            REGQ& = REGQ& + 1
            Call GRAREG("LIMULTI", String$(128, "="), REGQ&)
          End If
          '
          For II% = 1 To CONTAMUL%
            '
            If II% = 1 Then
              XX$ = REGBLAN$("LIMULTI")
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
            End If
            '
            If CANT(II%) > 0.00005 Then  ' SUPRIME DESARMADO
              XX$ = REGBLAN$("LIMULTI")
              Call REPLA(XX$, MKI$(II% - 1), 1, 2)
              Call REPLA(XX$, CODEXT$(CIK%(II%)), 3, 16)
              Call REPLA(XX$, DESCRIT$(CIK%(II%)), 17 + 2 * NIVE%(II%), 48)
              Call REPLA(XX$, CSTRING$(MKS$(CANT(II%)), 4, 10, 5, 2), 75, 10)
              Call REPLA(XX$, CSTRING$(MKI$(PAD%(II%)), 1, 6, 0, 2), 85, 6)
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
            End If
            '
            If II% = 1 Then
              XX$ = REGBLAN$("LIMULTI")
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
            End If
            '
          Next II%
          '
        End If
      End If
      U& = U& + 1
      MasterTemp.MoveNext
    Loop
    '
    Call SETULTREG("LIMULTI", REGQ&)
    Call CIERRARCH("LIMULTI")
    '
    Call HEADERS("LIMULTI", "")
    Screen.MousePointer = 1
    Call FINAL("*LIMULTI")
    '
End Sub
'

Sub LIMULIN()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
       Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    j& = 0
    Set MasterTemp = Articulos.OpenRecordset("SELECT * FROM Master WHERE CODIGO>='" & COD1$ & "' AND CODIGO<='" & COD2$ & "' ORDER BY CODIGO ASC;", dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    MasterTemp.MoveLast
    If Err Then Exit Sub
    '
    CAREC% = MasterTemp.RecordCount
    If CAREC% < 1 Then Exit Sub
    '
    FIN& = CAREC%
    MasterTemp.MoveFirst: U& = 1
    '
10  Call TRACE(20, U&, FIN&)
    CI% = MasterTemp!CODINT
    If STATUIT%(CI%) = 1 Or CI1% = CI2% Then
      PORMILLAR% = 0
      Call EXPLOCOM(CI%, 1)
      If CONTAMUL% > 0 Then
        '
        ACTIX% = 0
        For K% = 1 To CONTAMUL%
          CIX1% = CIK%(K%)
          CAN = CANT(K%)
          PD% = 0
          If CIX1% > 0 Then
            If CIX1% <= NUMAS% Then
              If Abs(CAN) > 0.0005 Then
                UM$ = UNIMED$(CIX1%)
                Z$ = REGBLAN$("LIESTRU")
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                Call REPLA(Z$, MKI$(CIX1%), 3, 2)
                Call REPLA(Z$, UM$, 5, 2)
                Call REPLA(Z$, MKS$(CAN), 7, 4)
                Call REPLA(Z$, Chr$(PD%), 11, 1)
                '
                j& = j& + 1
                Call GRAREG("LIESTRU", Z$, j&)
                ACTIX% = 1
                '
              End If
            End If
          End If
        Next K%
        '
        If ACTIX% > 0 Then
          Z$ = REGBLAN$("LIESTRU")
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          j& = j& + 1
          Call GRAREG("LIESTRU", Z$, j&)
        End If
        '
      End If
    End If
    '
    On Error Resume Next
    MasterTemp.MoveNext: U& = U& + 1
    If Err < 1 Then
      If U& <= CAREC% Then
        GoTo 10
      End If
    End If
    '
    Call SETULTREG("LIESTRU", j&)
    Call CIERRARCH("LIESTRU")
    '
    Call HEADERS("LIESTRU", "")
    Screen.MousePointer = 1
    Call FINAL("*LICOMPRA")
    '
End Sub
'

Sub IMPORFORMUL()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   If COMALTER%("Forma de Importación:\\Completa\Solo Fórmulas") = 2 Then GoTo 51
   '
   Label58.Visible = True
   IKK& = 0
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 15
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
10   If EOF(N1%) Then
       Close #N1%
       GoTo 15
     End If
     '
     Line Input #N1%, AA$
     IKK& = IKK& + 1: Label58.Caption = TRIM$(Str$(IKK&))
     Label58.Refresh
     COD$ = TRIM$(Mid$(AA$, 1, 15))
     DEX$ = TRIM$(Mid$(AA$, 21, 55))
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
11   COD$ = TRIM$(Mid$(AA$, 76, 15))
     DEX$ = TRIM$(Mid$(AA$, 92, 51))
     UNI$ = "U."
     COSUNIX$ = TRIM$(Mid$(AA$, 149, 10))
     Call REPLACAR(COSUNIX$, ",", ".")
     COU = Val(COSUNIX$)
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 12
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
12   REGMA& = KKI% + 1
     REMAS$ = REGLEIDO$("MASTER", REGMA&)
     If CVS(Mid$(REMAS$, 111, 4)) < 0.0001 Then
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
     End If
     '
     GoTo 10
     '
15 Next UI%
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 65
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 15))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR1$ = TRIM$(Mid$(AA$, 76, 15))
     If TRIM$(CODPA1R$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 143, 6))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR1$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 Next UI%
   '
   PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   j& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub
'

Private Sub Text9_GotFocus(Index As Integer)
   Text9(Index).SelStart = 0
   Text9(Index).SelLength = Len(Text9(Index).TEXT)
End Sub

Sub IMPORARCON()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   If COMALTER%("Forma de Importación:\\Completa\Solo Fórmulas") = 2 Then GoTo 51
   '
   Label58.Visible = True
   IKK& = 0
   '
     ARIMPO$ = "A:\MAESTRO.PRN"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
10   If EOF(N1%) Then
       Close #N1%
       GoTo 19
     End If
     '
     Line Input #N1%, AA$
     IKK& = IKK& + 1: Label58.Caption = TRIM$(Str$(IKK&))
     Label58.Refresh
     COD$ = TRIM$(Mid$(AA$, 1, 12))
     DEX$ = TRIM$(Mid$(AA$, 13, 41))
     UNI$ = TRIM$(Mid$(AA$, 64, 4))
     COSUNIX$ = TRIM$(Mid$(AA$, 83, 10))
     COSUN$ = REPLACAR(COSUNIX$, ",", ".")
     COU = Val(COSUN$)
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = String$(128, 0)
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
11 GoTo 10
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   ARIMPO$ = "A:\FORMULAS.PRN"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 13))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR1$ = TRIM$(Mid$(AA$, 14, 15))
     If TRIM$(CODPAR1$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 34, 10))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR1$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   j& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub

 Sub RECACOSVIE()
     '
     Screen.MousePointer = 11
     QQQ% = 1024
     Dim DP%(16384), DQ%(16384)
     Dim COMAQ$(1024), CUHOR(1024)
     '
     CAMAQ% = ULTREG&("PADMAQ")
     For UI& = 1 To ULTREG&("PADMAQ")
        R0$ = REGLEIDO$("PADMAQ", UI&)
        COMAQ$(UI&) = TRIM$(Left$(R0$, 6))
        CUHOR(UI&) = CVS(Mid$(R0$, 93, 4))
     Next UI&
     Call CIERRARCH("PADMAQ")
     '
62   HOI% = HOY(HO$)
AAA = CODINT%("12710001")
     NVI% = NUINV%
     NMM% = NUMAS%
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     '
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        R0$ = REGLEIDO$("ESTRUNUE", I&)
        R2$ = Mid$(R0$, 3, 120)
        CI% = CVI(Left$(R0$, 2))
        If CI% > 0 Then
            If CI% <= NMM% Then
                For K% = 1 To 120 Step 8
                    DPSX$ = Mid$(R2$, K%, 8)
                    If CVI(Left$(DPSX$, 2)) < 1 Or CVI(Left$(DPSX$, 2)) > NMM% Then GoTo 105
                    If CVS(Mid$(DPSX$, 3)) >= 0.0005 Then
                        DP%(CI%) = 1
                        DQ%(CI%) = 1
                        GoTo 105
                    End If
                Next K%
            End If
        End If
105  Next I&
     '
110  N9% = FILEOPEN%(LU$ + "COCALCU.DAT", 0, 128)
     'OPEN "R",#N9%,LU$+"COCALCU.DAT",128
     'FIELD #N9%,2 AS Q1$,2 AS Q3$,2 AS Q4$,4 AS Q11$,4 AS Q12$,4 AS Q13$,4 AS Q21$,4 AS Q22$,4 AS Q23$,4 AS Q31$,4 AS Q32$,4 AS Q33$,4 AS Q41$,4 AS Q42$,4 AS Q43$,74 AS Q99$
     Dim Q0 As String * 128
     N19% = FILEOPEN%(LU$ + "MODIREC.DAT", 0, 128)
     Dim Q9 As String * 128
     For II% = 1 To NMM%
        Get #N9%, II%, Q0$
        LSet Q0$ = String$(128, 0)
        Put #N9%, II%, Q0$
        Get #N19%, II%, Q9$
        LSet Q9$ = String$(127, 0) + Chr$(1)
        Put #N19%, II%, Q9$
     Next II%
115  CANA& = 0

     ' CARGANDO COSTOS DE SERVICIOS SUMINISTRADOS POR TERCEROS
     N6% = FILEOPEN(LU$ + "LPREPRO.DAT", 0, 48)
     Dim P0 As String * 48
     FIN& = ULTREG&("LPREPRO")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        P0$ = REGLEIDO$("LPREPRO", I&)
125     CI% = CVI(Mid$(P0$, 3, 2))
        COU = CVS(Mid$(P0$, 37, 4))
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N9%, CI%, Q0$
                RQ0$ = Q0$
                V31 = CVS(Mid$(Q0$, 31, 4)) + COU
                V33 = CVS(Mid$(Q0$, 39, 4)) + COU
                V41 = CVS(Mid$(Q0$, 43, 4)) + COU
                V43 = CVS(Mid$(Q0$, 51, 4)) + COU
                Call REPLA(RQ0$, MKS$(V31), 31, 4)
                Call REPLA(RQ0$, MKS$(V33), 39, 4)
                Call REPLA(RQ0$, MKS$(V41), 43, 4)
                Call REPLA(RQ0$, MKS$(V43), 51, 4)
                '
                Q99$ = Mid$(RQ0$, 55, 74)
                POIN$ = ""
                For U% = 1 To 74 Step 2
                    U1% = CVI(Mid$(Q99$, U%, 2))
                    If U1% > 0 Then
                        POIN$ = POIN$ + MKI$(U1%)
                    End If
                Next U%
                POIN$ = POIN$ + MKI$(KI%) + String$(72, 0)
                Call REPLA(RQ0$, POIN$, 55, 74)
                '
                LSet Q0$ = RQ0$
                Put #N9%, CI%, Q0$
            End If
        End If
     Next I&
     '
     'CARGANDO COSTOS DE MANO DE OBRA DIRECTA
     FIN& = ULTREG&("SECOPER")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        SCP$ = REGLEIDO$("SECOPER", I&)
        CI% = CVI(Left$(SCP$, 2))
        MAQUI$ = TRIM$(Mid$(SCP$, 53, 6))
        CUOTHOR = 0
        'If TRIM$(MAQUI$) <> "" Then
           For KK% = 1 To CAMAQ%
              If COMAQ$(KK%) = MAQUI$ Then
                 CUOTHOR = CUHOR(KK%)
              End If
           Next KK%
        'End If
        '
        LOTESTA = LOTESTAN(CI%)
        If LOTESTA < 1 Then LOTESTA = 1
        CALOTES = 1
        MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
        MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
        MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
        MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
        MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
        MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
        HORAPROD = MINLOTE / 60
        CAOPS = CVS(Mid$(SCP$, 83, 4))
        CATOP% = Asc(Mid$(SCP$, 87, 1))
        CUOTOPE = VALHORA(CATOP%)
        COU = (CUOTHOR + CUOTOPE * CAOPS) * HORAPROD / LOTESTA
        '
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N19%, CI%, Q9$
                RQ0$ = Q9$
                V61 = CVS(Mid$(Q9$, 3, 4)) + COU
                V63 = CVS(Mid$(Q9$, 11, 4)) + COU
                Call REPLA(RQ0$, MKS$(V61), 3, 4)
                Call REPLA(RQ0$, MKS$(V63), 11, 4)
                LSet Q9$ = RQ0$
                Put #N19%, CI%, Q9$
            End If
        End If
     Next I&
     '
     ' TRANSFIRIENDO COSTOS DE REPOSICION DE ITEMS SIN EXPLOSION
     FIN& = NUMAS%
     Screen.MousePointer = 11
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        M0$ = REGLEIDO$("MASTER", I& + 1)
        CI% = CVI(Left$(M0$, 2))
        UNI$ = Mid$(M0$, 109, 2)
        Get #N9%, CI%, Q0$
        RQ0$ = Q0$
        Call REPLA(RQ0$, MKI$(CI%), 1, 2)
        Call REPLA(RQ0$, UNI$, 3, 2)
        Call REPLA(RQ0$, "$ ", 5, 2)
        LSet Q0$ = RQ0$
        Put #N9%, CI%, Q0$
        '
        If DP%(CI%) = 0 Then
            COU = CVS(Mid$(M0$, 111, 4))
            Get #N19%, CI%, Q9$
            If CVS(Mid$(Q9$, 3, 4)) > 0.000005 Then
               COU = 0 ' si tiene M.Obra no puede tener Costo de Materiales
            End If
            MONE$ = "$"
            If Asc(Mid$(M0$, 95, 1)) = 2 Then MONE$ = "U$"
            If Asc(Mid$(M0$, 95, 1)) = 3 Then MONE$ = "DM"
            '
            Call REPLA(RQ0$, MONE$, 5, 2)
            If UNI$ = "U." Then
                 Call REPLA(RQ0$, MKS$(COU), 19, 4)
                 Call REPLA(RQ0$, MKS$(COU), 27, 4)
               Else
                 Call REPLA(RQ0$, MKS$(COU), 7, 4)
                 Call REPLA(RQ0$, MKS$(COU), 15, 4)
            End If
            Call REPLA(RQ0$, MKS$(COU), 43, 4)
            Call REPLA(RQ0$, MKS$(COU), 51, 4)
            LSet Q0$ = RQ0$
            Put #N9%, CI%, Q0$

            If COU > 0.00005 Then
                CANA& = CANA& + 1
                RW0$ = String$(96, 0)
                Call REPLA(RW0$, MKI$(CI%), 1, 2)
                CIXI% = CI%
                Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                CIXI% = CI%
                Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                Call REPLA(RW0$, MONE$, 67, 2)
                Call REPLA(RW0$, MKS$(COU), 69, 4)
                Call REPLA(RW0$, MKS$(1), 73, 4)
                Call REPLA(RW0$, MKS$(COU), 77, 4)
                Call REPLA(RW0$, MKS$(0), 81, 4)
                Call REPLA(RW0$, MKS$(0), 85, 4)
                Call REPLA(RW0$, MKS$(COU), 89, 4)
                Call REPLA(RW0$, MKS$(0), 93, 4)
                '
                Call GRAREG("ANACOS", RW0$, CANA&)
            End If

130         If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                POIN$ = Mid$(Q0$, 55, 74)
                For U% = 1 To Len(POIN$) Step 2
                    POINI% = CVI(Mid$(POIN$, U%, 2))
                    If POINI% > 0 Then
                        CIXI% = CI%
                        Get #N6%, POINI%, P0$
                        DE$ = TRIM$(Mid$(P0$, 5, 30))
                        MN$ = Mid$(P0$, 35, 2)
                        COU = CVS(Mid$(P0$, 37, 4))
                        CANA& = CANA& + 1
                        '
                        RW0$ = String$(96, 0)
                        Call REPLA(RW0$, MKI$(CI%), 1, 2)
                        CIXI% = CI%
                        Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                        CIXI% = CI%
                        Call REPLA(RW0$, DE$, 19, 48)
                        Call REPLA(RW0$, MN$, 67, 2)
                        Call REPLA(RW0$, MKS$(COU), 69, 4)
                        Call REPLA(RW0$, MKS$(1), 73, 4)
                        Call REPLA(RW0$, MKS$(0), 77, 4)
                        Call REPLA(RW0$, MKS$(0), 81, 4)
                        Call REPLA(RW0$, MKS$(COU), 85, 4)
                        Call REPLA(RW0$, MKS$(COU), 89, 4)
                        Call REPLA(RW0$, MKS$(0), 93, 4)
                        Call GRAREG("ANACOS", RW0$, CANA&)
                        '
                    End If
                Next U%
            End If
        End If
     Next I&
     ' RE-CALCULANDO COSTOS STANDARD DE ITEMS SEGUN EXPLOSION

140  RECA% = 0
     FIN& = ULTREG&("ESTRUNUE")
     Screen.MousePointer = 11
     For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         R0$ = Left$(REGLEIDO$("ESTRUNUE", I&), 122)
         If I& <= ULTREG&("ESTRUDOS") Then
              R0$ = R0$ + REGLEIDO$("ESTRUDOS", I&)
            Else
              R0$ = R0$ + String$(256, 0)
         End If
         CI% = CVI(Left$(R0$, 2))
         DPS$ = Mid$(R0$, 3)
         If CI% < 1 Or CI% > NUMAS% Then GoTo 199

142      If DP%(CI%) = 0 Then GoTo 199                          ' ya fue calculado

         CIXI% = CI%
         CABAS = 1: SUMP = 0: SUMF = 0: SUS3 = 0: SUMO = 0
145      For K% = 1 To 376 Step 8
             DPSX$ = Mid$(DPS$, K%, 8)
             CI1% = CVI(DPSX$)
             If CI1% < 1 Or CI1% > NUMAS% Then GoTo 149    ' terminar
             '
             If CI1% = CI% Then
                If COMALTER%("Error de Recursividad en la Fórmula\Correspondiente a " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".\\Cancelar\Continuar") <> 2 Then
                   Call GRACONTROL("COSTOS", "FECALCOS", "00/00/00")
                   Call GRACONTROL("COSTOS", "HOCALCOS", "00:00:00")
                   FORCOSTO.Label34.Caption = " " + HOS$
                   FORCOSTO.Label34.Refresh
                   FORCOSTO.Label35.Caption = " " + Time$
                   FORCOSTO.Label35.Refresh
                   GoTo 3800
                End If
             End If
             '
             If DP%(CI1%) > 0 Then GoTo 199                ' aun no fue calculado
             CANT7 = CABAS * CVS(Mid$(DPSX$, 3))
             Get #N9%, CI1%, Q0$
             Get #N19%, CI1%, Q9$
             SUMP = SUMP + CANT7 * CVS(Mid$(Q0$, 15, 4))
             SUMF = SUMF + CANT7 * CVS(Mid$(Q0$, 27, 4))
             SUS3 = SUS3 + CANT7 * CVS(Mid$(Q0$, 39, 4))
             MOOX = CVS(Mid$(Q9$, 11, 4))
             SUMO = SUMO + CANT7 * CVS(Mid$(Q9$, 11, 4))
'BBBB$ = CODEXT$(CI1%)
         Next K%
         '
149      Get #N9%, CI%, Q0$
         RQ0$ = Q0$
         Get #N19%, CI%, Q9$
         RQ9$ = Q9$
         '
         V12 = CVS(Mid$(Q0$, 11, 4)) + SUMP
         V13 = CVS(Mid$(Q0$, 15, 4)) + SUMP
         V22 = CVS(Mid$(Q0$, 23, 4)) + SUMF
         V23 = CVS(Mid$(Q0$, 27, 4)) + SUMF
         V32 = CVS(Mid$(Q0$, 35, 4)) + SUS3
         V33 = CVS(Mid$(Q0$, 39, 4)) + SUS3
         V42 = CVS(Mid$(Q0$, 47, 4)) + SUMP + SUMF + SUS3
         V43 = CVS(Mid$(Q0$, 51, 4)) + SUMP + SUMF + SUS3
         V62 = CVS(Mid$(Q9$, 7, 4)) + SUMO
         V63 = CVS(Mid$(Q9$, 11, 4)) + SUMO
         '
         Call REPLA(RQ0$, MKS$(V12), 11, 4)
         Call REPLA(RQ0$, MKS$(V13), 15, 4)
         Call REPLA(RQ0$, MKS$(V22), 23, 4)
         Call REPLA(RQ0$, MKS$(V23), 27, 4)
         Call REPLA(RQ0$, MKS$(V32), 35, 4)
         Call REPLA(RQ0$, MKS$(V33), 39, 4)
         Call REPLA(RQ0$, MKS$(V42), 47, 4)
         Call REPLA(RQ0$, MKS$(V43), 51, 4)
         '
         Call REPLA(RQ9$, MKS$(V62), 7, 4)
         Call REPLA(RQ9$, MKS$(V63), 11, 4)
         '
         LSet Q0$ = RQ0$
         Put #N9%, CI%, Q0$
         '
         LSet Q9$ = RQ9$
         Put #N19%, CI%, Q9$
         QBAS = 1: CIX0% = CI%
         Call EXPLOMUL(CI%)
         Screen.MousePointer = 11
         For j% = 1 To CONTAMUL%
             CII% = CIK%(j%)
             If CII% > 0 Then
                If CII% <= NUMAS% Then
                   Get #N9%, CII%, Q0$
                   If DQ%(CII%) = 0 Then
                      COU = CVS(Mid$(Q0$, 7, 4)) + CVS(Mid$(Q0$, 19, 4))
                      CANA& = CANA& + 1
                      RW0$ = String$(96, 0)
                      Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                      CIXI% = CII%
                      Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                      CIXI% = CII%
                      Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                      Call REPLA(RW0$, "$ ", 67, 2)
                      Call REPLA(RW0$, MKS$(COU), 69, 4)
                      Call REPLA(RW0$, MKS$(CANT(j%)), 73, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(j%)), 77, 4)
                      Call REPLA(RW0$, MKS$(0), 81, 4)
                      Call REPLA(RW0$, MKS$(0), 85, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(j%)), 89, 4)
                      Call REPLA(RW0$, MKS$(0), 93, 4)
                      Call GRAREG("ANACOS", RW0$, CANA&)
                   End If
                   '
                   If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                      POIN$ = Mid$(Q0$, 55, 74)
                      For U% = 1 To Len(POIN$) Step 2
                         POINI% = CVI(Mid$(POIN$, U%, 2))
                         If POINI% > 0 Then
                            CIXI% = CI%
                            Get #N6%, POINI%, P0$
                            DE$ = TRIM$(Mid$(P0$, 5, 30))
                            MN$ = Mid$(P0$, 35, 2)
                            COU = CVS(Mid$(P0$, 37, 4))
                            CANA& = CANA& + 1
                            '
                            RW0$ = String$(96, 0)
                            Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                            CIXI% = CII%
                            Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                            CIXI% = CII%
                            Call REPLA(RW0$, DE$, 19, 48)
                            Call REPLA(RW0$, MN$, 67, 2)
                            Call REPLA(RW0$, MKS$(COU), 69, 4)
                            Call REPLA(RW0$, MKS$(CANT(j%)), 73, 4)
                            Call REPLA(RW0$, MKS$(0), 77, 4)
                            Call REPLA(RW0$, MKS$(0), 81, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(j%)), 85, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(j%)), 89, 4)
                            Call REPLA(RW0$, MKS$(0), 93, 4)
                            Call GRAREG("ANACOS", RW0$, CANA&)
                            '
                         End If
                      Next U%
                   End If
                End If
             End If
         Next j%
         '
         DP%(CIX0%) = 0
         RECA% = 1
         '
199  Next I&
     If RECA% > 0 Then GoTo 140

210  JJ& = 0
     For I& = 1 To NUINV%
        I0$ = REGLEIDO$("INVERT", I& + 1)
        COD$ = Left$(I0$, 16): CI% = CVI(Mid$(I0$, 17, 2))
        If Asc(Mid$(COD$, 16, 1)) <> 96 Then
            If CI% > 0 Then
                If CI% <= NUMAS% Then
                    Get #N9%, CI%, Q0$
                    Z$ = Left$(Q0$, 64)
                    JJ& = JJ& + 1
                    Call GRAREG("LICOCAL", Z$, JJ&)
                End If
            End If
        End If
     Next I&
     '
     Call SETULTREG("LICOCAL", JJ&)
     Call SETULTREG("ANACOS", CANA&)
     '
     FEX = HOY(HOS$)
     Call GRACONTROL("COSTOS", "FECALCOS", HOS$)
     Call GRACONTROL("COSTOS", "HOCALCOS", Time$)
     '
     FORCOSTO.Label34.Caption = " " + HOS$
     FORCOSTO.Label34.Refresh
     FORCOSTO.Label35.Caption = " " + Time$
     FORCOSTO.Label35.Refresh
     '
3800 Call CIERRARCH("*.*")
     Close #N9%
     Close #N19%
3900 Screen.MousePointer = 1
     '
End Sub

Sub MUECOSTO()
   '
   RFF$ = RECFILT$("COCALCU", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
      I& = CVS(RFF$)
      Q0$ = REGLEIDO$("COCALCU", I&)
      Q9$ = REGLEIDO$("MODIREC", I&)
      If CVI(Left$(Q0$, 2)) = CI% Then
         Label90.Caption = TRIM$(CSTRING$(Mid$(Q0$, 7, 4), 3, 12, 4, 2))
         Label90.Refresh
         Label82.Caption = TRIM$(CSTRING$(Mid$(Q0$, 11, 4), 3, 12, 4, 2))
         Label82.Refresh
         Label88.Caption = TRIM$(CSTRING$(Mid$(Q0$, 19, 4), 3, 12, 4, 2))
         Label88.Refresh
         Label83.Caption = TRIM$(CSTRING$(Mid$(Q0$, 23, 4), 3, 12, 4, 2))
         Label83.Refresh
         Label86.Caption = TRIM$(CSTRING$(Mid$(Q0$, 31, 4), 3, 12, 4, 2))
         Label86.Refresh
         Label84.Caption = TRIM$(CSTRING$(Mid$(Q0$, 35, 4), 3, 12, 4, 2))
         Label84.Refresh
         Label75.Caption = TRIM$(CSTRING$(Mid$(Q9$, 3, 4), 3, 12, 4, 2))
         Label75.Refresh
         Label72.Caption = TRIM$(CSTRING$(Mid$(Q9$, 7, 4), 3, 12, 4, 2))
         Label72.Refresh
         '
         SUMA1 = 0
         SUMA1 = SUMA1 + Val(Label90.Caption)
         SUMA1 = SUMA1 + Val(Label88.Caption)
         SUMA1 = SUMA1 + Val(Label86.Caption)
         SUMA1 = SUMA1 + Val(Label82.Caption)
         SUMA1 = SUMA1 + Val(Label83.Caption)
         SUMA1 = SUMA1 + Val(Label84.Caption)
         Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, 4, 2))
         Label85.Refresh
         '
         SUMA2 = 0
         SUMA2 = SUMA2 + Val(Label75.Caption)
         SUMA2 = SUMA2 + Val(Label72.Caption)
         Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, 4, 2))
         Label80.Refresh
         '
         SUMA3 = SUMA1
         SUMA3 = SUMA3 + Val(Label80.Caption)
         Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, 4, 2))
         Label70.Refresh
         '
         SUMA4 = SUMA3 * (Val(Text4.TEXT)) / 100
         Label68.Caption = TRIM$(CSTRING$(MKS$(SUMA4), 3, 12, 4, 2))
         Label68.Refresh
         '
         SUMA5 = SUMA3 + SUMA4
         Label65.Caption = TRIM$(CSTRING$(MKS$(SUMA5), 3, 12, 4, 2))
         Label65.Refresh
         '
      End If
   End If
End Sub

Sub CARCOSTO()
   '
   Static DP%(32767), FESTRU#
   '
   Screen.MousePointer = 11
   Call FRESHALL
   Call BLANFORCOS
   GoSub 90
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   CICON% = CI%
   '
   NVI% = NUINV%
   NMM% = NUMAS%
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   'FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   'If FESTRUC# <> FESTRU# Then
   '  FESTRU# = FESTRUC#
   '  FIN& = ULTREG&("ESTRUNUE")
   '  For I& = 1 To FIN&
   '    Call TRACE(20, I&, FIN&)
   '    R0$ = REGLEIDO$("ESTRUNUE", I&)
   '    CI1% = CVI(Left$(R0$, 2))
   '    If CI1% > 0 Then
   '      If CI1% <= NMM% Then
   '        If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
   '          DP%(CI1%) = 1
   '        End If
   '      End If
   '    End If
   '  Next I&
   'End If
   '
   Call EXPLOMUL(CI%)
   SUMPPROP = 0: SUMPREC = 0
   SUPIPROP = 0: SUPIREC = 0
   SUSEPROP = 0: SUSEREC = 0
   '
   For KKI% = 1 To CONTAMUL%
     For KKJ% = KKI% + 1 To CONTAMUL%
       If CIK%(KKJ%) = CIK%(KKI%) Then
         CANT(KKI%) = CANT(KKI%) + CANT(KKJ%)
         CANT(KKJ%) = 0
       End If
     Next KKJ%
   Next KKI%
   '
   ' costo de materias primas y piezas de terceros
   '
   ANACOSMAT.List1.Clear
   SUPRI = 0: SUPIE = 0: SUSER = 0
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
     If Abs(CANT(KKI%)) < 0.000005 Then GoTo 19
     If KKI% >= CONTAMUL% Or NIVE%(KKI% + 1) <= NIVE%(KKI%) Then  ' NO TIENE ESTRUCTURA
'     II& = KKI%: Call TRACE(20, II&, FIN&)
'     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
15      MONEI% = MONECOSTO%(CIK%(KKI%))
        DEMONE$ = DEMONECO$(MONEI%)
        COSUDOL = COSUNI(CIK%(KKI%))
        If MONEI% > 1 Then
            COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
          Else
            FEBAS% = FECOSTO%(CIK%(KKI%))
            FEACT% = HOY(HOS$)
            COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
        End If
        '
        Z$ = Space$(256)
        Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
        DESCRI$ = DESCRIT0$(CIK%(KKI%))
        Call REPLA(Z$, DESCRI$, 17, 23)
        Call REPLA(Z$, DESCRI$, 129, 40)
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
        Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
        Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
        Call REPLA(Z$, DEMONE$, 179, 10)
        Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
        UNIME$ = UCase$(UNIMED$(CIK%(KKI%)))
        If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
             If KKI% = 1 Then
                  SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                Else
                  SUPIREC = SUPIREC + CANT(KKI%) * COSUN
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 65, 10)
             SUPIE = SUPIE + CANT(KKI%) * COSUN
          Else
             If KKI% = 1 Then
                  SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                Else
                  SUMPREC = SUMPREC + COSUN * CANT(KKI%)
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 55, 10)
             SUPRI = SUPRI + CANT(KKI%) * COSUN
        End If
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
        ANACOSMAT.List1.AddItem Z$
        '
      Else
        '
        If COSUNI(CIK%(KKI%)) > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
             If PROVINT%(PHB%) = 0 Then
               MONEI% = MONECOSTO%(CIK%(KKI%))
               DEMONE$ = DEMONECO$(MONEI%)
               COSUDOL = COSUNI(CIK%(KKI%))
               If MONEI% > 1 Then
                   COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                 Else
                   FEBAS% = FECOSTO%(CIK%(KKI%))
                   FEACT% = HOY(HOS$)
                   COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
               End If
               '
               Z$ = Space$(256)
               Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
               DESCRI$ = DESCRIT0$(CIK%(KKI%))
               Call REPLA(Z$, DESCRI$, 17, 23)
               Call REPLA(Z$, DESCRI$, 129, 40)
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
               Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
               Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
               Call REPLA(Z$, DEMONE$, 179, 10)
               Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
               If KKI% = 1 Then
                    SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                  Else
                    SUSEREC = SUSEREC + CANT(KKI%) * COSUN
               End If
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 75, 10)
               SUSER = SUSER + CANT(KKI%) * COSUN
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
               ANACOSMAT.List1.AddItem Z$
             End If
           End If
        End If
        '
     End If
19 Next KKI%
   '
   For KKI% = 1 To ANACOSMAT.List1.ListCount
22   TTXX$ = ANACOSMAT.List1.List(KKI% - 1)
     For KKJ% = KKI% + 1 To ANACOSMAT.List1.ListCount
       TTYY$ = ANACOSMAT.List1.List(KKJ% - 1)
       If "A" + TTXX$ > "A" + TTYY$ Then
         ANACOSMAT.List1.RemoveItem (KKJ% - 1)
         ANACOSMAT.List1.AddItem TTYY$, KKI% - 1
         GoTo 22
       End If
     Next KKJ%
   Next KKI%
   '
   If ANACOSMAT.List1.ListCount > 0 Then
      ANACOSMAT.List1.AddItem String$(128, "=")
      ANACOSMAT.Label1.Caption = TRIM$(CSTRING$(MKS$(SUPRI), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label4.Caption = TRIM$(CSTRING$(MKS$(SUPIE), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label3.Caption = TRIM$(CSTRING$(MKS$(SUSER), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label5.Caption = TRIM$(CSTRING$(MKS$(SUPRI + SUPIE + SUSER), 3, 10, DECICALCOS%, 2))
   End If
   '
   GoSub 90
   '
   ' costo de mano de obra directa
   '
   FIN& = CONTAMUL%
   SUMODPRO = 0: SUMODREC = 0
   SUMAQPRO = 0: SUMAQREC = 0
   '
   ANACOSMOD.List1.Clear
   SUTIE = 0: SUMOD = 0: SUMAQ = 0
   ICX% = 0
   '
   For UI& = 1 To CONTAMUL%
      '
      CIXI% = CIK%(UI&)
      CODEX$ = CODEXT$(CIXI%)
      CACOM = CANT(UI&)
      '
      SQLX$ = "SELECT * FROM Secoper"
      SQLX$ = SQLX$ & " WHERE CodIConj = " & CIXI%
      SQLX$ = SQLX$ & " AND StatOper >0 "
      Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      IOP% = 0
      On Error Resume Next
      With SecopTemp
        .MoveFirst
        If Err Then GoTo 25
        '
        Do Until .EOF = True
          If IsNull(!AltOper) = True Or !AltOper < 1 Then
             NOPY% = !NumeOper
             DOPE$ = !DescOper
             LOTESTA = LOTESTAN(CIXI%)
                If LOTESTA < 1 Then LOTESTA = 1
             CICLOHE = !PieCiclo
                If CICLOHE < 1 Then CICLOHE = 1
             MINPREP = 60 * !HorsPrep
             MINFIJO = 60 * !HorsFijo
             MINVARI = LOTESTA * !MinuStan
             MINHERR = LOTESTA * 60 * !HorsMHer / CICLOHE
             MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
             MINLOTE = MINFIJO + MINVARI + MINHERR + MINDEMO
             HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
             '
             CAOPS = !CantiOps
                 CATOP% = !CategOps ' Asc(Mid$(SCP$, 87, 1))
             EQUIPO$ = !CodMaq0
             '
             CUOTOPE = VALHORA(CATOP%)
             CUOTMAQ = CUHORMAQ(EQUIPO$)
             '
             CAOPREP = CANOPREP(CIXI%, NOPY%)
             If CAOPREP < 1 Then CAOPREP = CAOPS
             CTOPREP% = CATOPREP%(CIXI%, NOPY%)
             If CTOPREP% < 1 Then CTOPREP% = CATOP%
             CUOTOPREP = VALHORA(CTOPREP%)
             '
             COSMOD = CUOTOPE * CAOPS * HORAPROD + CUOTOPREP * CAOPREP * MINPREP / 60 / LOTESTA
             COSMAQ = CUOTMAQ * (HORAPROD + MINPREP / 60 / LOTESTA)
             '
             TERCOP% = 0: COSTER = 0
             MONEI% = !MonePrec
             COSTER = CACOM * !PreUnid * TICAMONE(MONEI%)
             If MONEI% <= 1 Then
                   'FEBAS% = CVI(Mid$(SCT$, 17, 2))
                   '         FEACT% = HOY(HOS$)
                   COSTER = CACOM * !PreUnid * IAJUINF#(FEBAS%, FEACT%)
             End If
             If COSTER >= 0.00005 Then
               If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                 COSMOD = COSTER
                 COSMAQ = 0
                 HORAPROD = 0
                 TERCOP% = 1
               End If
             End If
             '
             SUTIE = SUTIE + HORAPROD
             SUMOD = SUMOD + COSMOD
             SUMAQ = SUMAQ + COSMAQ
             '
             If UI& = 1 Then
                 SUMODPRO = SUMODPRO + COSMOD
                 SUMAQPRO = SUMAQPRO + COSMAQ
               Else
                 SUMODREC = SUMODREC + COSMOD
                 SUMAQREC = SUMAQREC + COSMAQ
             End If
             '
             IOP% = IOP% + 1: ICX% = ICX% + 1
             Z$ = Space$(256)
             If IOP% = 1 Then
               If ICX% > 1 Then
                 ANACOSMOD.List1.AddItem String$(256, "-")
               End If
               Call REPLA(Z$, CODEXT$(CIXI%), 1, 15)
               DESCRI$ = DESCRIT0$(CIXI%)
               Call REPLA(Z$, DESCRI$, 17, 21)
               Call REPLA(Z$, DESCRI$, 129, 40)
             End If
             Call REPLA(Z$, DOPE$, 39, 12)
             Call REPLA(Z$, DOPE$, 169, 30)
             Call REPLA(Z$, CSTRING$(MKS$(CAOPS), 3, 5, 2, 2), 199, 5)
             Call REPLA(Z$, HORMINSEMIL$(3600 * HORAPROD), 51, 14)
             If TERCOP% = 1 Then
               Call REPLA(Z$, "  (3ros)", 55, 10)
             End If
             Call REPLA(Z$, CSTRING$(MKS$(COSMOD), 3, 10, DECICALCOS%, 2), 65, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COSMAQ), 3, 10, DECICALCOS%, 2), 75, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COSMOD + COSMAQ), 3, 10, DECICALCOS%, 2), 85, 10)
             If COSTER * TERCOP% > 0 Then
               Call REPLA(Z$, CSTRING$(MKS$(COSTER / CACOM), 3, 10, DECICALCOS%, 2), 204, 10)
             End If
             ANACOSMOD.List1.AddItem Z$
             '
          End If
          .MoveNext
        Loop
        '
      End With
      '
25 Next UI&
   '
   If ANACOSMOD.List1.ListCount > 0 Then
      ANACOSMOD.List1.AddItem String$(256, "=")
      ANACOSMOD.Label3.Caption = HORMINSEMIL$(3600 * SUTIE)
      ANACOSMOD.Label1.Caption = TRIM$(CSTRING$(MKS$(SUMOD), 3, 10, DECICALCOS%, 2))
      ANACOSMOD.Label4.Caption = TRIM$(CSTRING$(MKS$(SUMAQ), 3, 10, DECICALCOS%, 2))
      ANACOSMOD.Label5.Caption = TRIM$(CSTRING$(MKS$(SUMOD + SUMAQ), 3, 10, DECICALCOS%, 2))
   End If
   '
   GoSub 90
   '
   '
   YACAR%(4) = 1
   MODIFIC%(4) = 0
   Screen.MousePointer = 1
   '
   If ORILLAMA$ = "MAT" Then
     If REGILLAMA% >= 0 Then
       If REGILLAMA% < ANACOSMAT.List1.ListCount Then
         ANACOSMAT.List1.ListIndex = REGILLAMA%
         On Error Resume Next
           ANACOSMAT.List1.TopIndex = TOPELLAMA%
         On Error GoTo 0
       End If
     End If
     ORILLAMA$ = ""
     REGILLAMA% = -1
     TOPELLAMA% = 0
     LLAMACOSTO$ = ""
     ANACOSMAT.Show 1
   End If
   '
   If ORILLAMA$ = "MOD" Then
     If REGILLAMA% >= 0 Then
       If REGILLAMA% < ANACOSMOD.List1.ListCount Then
         ANACOSMOD.List1.ListIndex = REGILLAMA%
         On Error Resume Next
           ANACOSMOD.List1.TopIndex = TOPELLAMA%
         On Error GoTo 0
       End If
     End If
     ORILLAMA$ = ""
     REGILLAMA% = -1
     TOPELLAMA% = 0
     LLAMACOSTO$ = ""
     ANACOSMOD.Show 1
   End If
   '
Exit Sub
   '
   '
   '
90 Label90.Caption = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   Label90.Refresh
   Label82.Caption = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   Label82.Refresh
   Label88.Caption = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   Label88.Refresh
   Label83.Caption = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   Label83.Refresh
   Label86.Caption = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   Label86.Refresh
   Label84.Caption = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   Label84.Refresh
   '
   Label75.Caption = TRIM$(CSTRING$(MKS$(SUMODPRO), 3, 12, DECICALCOS%, 2))
   Label75.Refresh
   Label72.Caption = TRIM$(CSTRING$(MKS$(SUMODREC), 3, 12, DECICALCOS%, 2))
   Label72.Refresh
   Label79.Caption = TRIM$(CSTRING$(MKS$(SUMAQPRO), 3, 12, DECICALCOS%, 2))
   Label79.Refresh
   Label74.Caption = TRIM$(CSTRING$(MKS$(SUMAQREC), 3, 12, DECICALCOS%, 2))
   Label74.Refresh
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(Label90.Caption)
   SUMA1 = SUMA1 + Val(Label88.Caption)
   SUMA1 = SUMA1 + Val(Label86.Caption)
   SUMA1 = SUMA1 + Val(Label82.Caption)
   SUMA1 = SUMA1 + Val(Label83.Caption)
   SUMA1 = SUMA1 + Val(Label84.Caption)
   Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, DECICALCOS%, 2))
   Label85.Refresh
   '
   SUMA2 = 0
   SUMA2 = SUMA2 + Val(Label75.Caption)
   SUMA2 = SUMA2 + Val(Label72.Caption)
   SUMA2 = SUMA2 + Val(Label79.Caption)
   SUMA2 = SUMA2 + Val(Label74.Caption)
   Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, DECICALCOS%, 2))
   Label80.Refresh
   '
   SUMA3 = SUMA1 + SUMA2
   Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, DECICALCOS%, 2))
   Label70.Refresh
   '
Return
'
End Sub

Sub CANACOSMAT(CIXX%)
   '
'Exit Sub
   Static ENCALI$, LINSEP$
   Static DP%(16384), FESTRU#
   '
   Screen.MousePointer = 11
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
     NVI% = NUINV%
     NMM% = NUMAS%
     FESTRU# = FESTRUC#
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = REGLEIDO$("ESTRUNUE", I&)
       CI1% = CVI(Left$(R0$, 2))
       If CI1% > 0 Then
         If CI1% <= NMM% Then
           DP%(CI1%) = 1
         End If
       End If
     Next I&
   End If
   '
   If ENCALI$ = "" Then
      ENCALI$ = "Codigo          Descripcion                                    Operacion                             $/Unid      Cant     Tot.Mat    Tot.Serv    Tot.Comp"
      LINSEP$ = String$(154, "-")
   End If
   '
   UREANA& = ULTREG&("ANACOS")
   Dim RECOSMAT$(4096)
   CACOM% = 0
   JJ& = 0
   '
   FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMAT.TXT", 2, 0)
   Call EXPLOMUL(CIXX%)
   For KKI% = 2 To CONTAMUL
     If DP%(CIK%(KKI%)) < 1 Then
       CODD$ = CODEXT$(CIK%(KKI%))
       DECO$ = DESCRIT$(CIK%(KKI%))
       DEOP$ = ""
       VUNI = COSUNI(CIK%(KKI%))
       MONEI% = MONECOSTO%(CIK%(KKI%))
       If MONEI% = 2 Then VUNI = VUNI * TICAMBIO
       CANX = CANT(KKI%)
       CMAT = CANX * VUNI
       '
       TEBUSCA$ = Space$(75)
       Call REPLA(TEBUSCA$, CODD$, 1, 16)
       Call REPLA(TEBUSCA$, DECO$, 17, 47)
       'Call REPLA(TEBUSCA$, DEOP$, 64, 32)
       Call REPLA(TEBUSCA$, CSTRING$(MKS$(VUNI), 3, 12, 4, 2), 64, 12)
       '
       For UJ% = 1 To CACOM%
         If Left$(RECOSMAT$(UJ%), 75) = TEBUSCA$ Then
           CANX = CANX + Val(Mid$(RECOSMAT$(UJ%), 76, 10))
           CMAT = CMAT + Val(Mid$(RECOSMAT$(UJ%), 86, 12))
           CSER = CSER + Val(Mid$(RECOSMAT$(UJ%), 98, 12))
           TENCU$ = RECOSMAT$(UJ%)
           Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 76, 10)
           Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 86, 12)
           Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 98, 12)
           RECOSMAT(UJ%) = TENCU$
           GoTo 9
         End If
       Next UJ%
       CACOM% = CACOM% + 1
       TENCU$ = Space$(154)
       Call REPLA(TENCU$, TEBUSCA$, 1, 75)
       Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 76, 10)
       Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 86, 12)
       Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 98, 12)
       RECOSMAT(CACOM%) = TENCU$
     End If
9  Next KKI%
   '
   For UI% = 1 To CACOM%
     For UJ% = UI% + 1 To CACOM%
       If Left$(RECOSMAT$(UI%), 16) > Left$(RECOSMAT$(UJ%), 16) Then
         TENCU$ = RECOSMAT$(UJ%)
         RECOSMAT$(UJ%) = RECOSMAT$(UI%)
         RECOSMAT$(UI%) = TENCU$
       End If
     Next UJ%
   Next UI%
   '
   SUMAPAR = 0: SUMATO = 0: SUMATER = 0: SUSERVI = 0
   Print #FISAL%, ENCALI$
   Print #FISAL%, LINSEP$
   For UI% = 1 To CACOM%
     Z$ = RECOSMAT$(UI%)
     SUMAPAR = SUMAPAR + Val(Mid$(Z$, 86, 12)) + Val(Mid$(Z$, 98, 12))
     SUMATO = SUMATO + Val(Mid$(Z$, 86, 12)) + Val(Mid$(Z$, 98, 12))
     SUMATER = SUMATER + Val(Mid$(Z$, 86, 12))
     SUSERVI = SUSERVI + Val(Mid$(Z$, 98, 12))
     If UI% = CACOM% Or Left$(RECOSMAT$(UI%), 16) <> Left$(RECOSMAT$(UI% + 1), 16) Then
       Call REPLA(Z$, CSTRING$(MKS$(SUMAPAR), 3, 12, 4, 2), 110, 12)
       SUMAPAR = 0
       ATXX% = 1
     End If
     Print #FISAL%, Z$
     If ATXX% = 1 Then Print #FISAL%, LINSEP$
   Next UI%
   '
   Z$ = Space$(154)
   Call REPLA(Z$, "S U B T O T A L   Materiales y Servicios", 1, 40)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATER), 3, 12, 4, 2), 86, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUSERVI), 3, 12, 4, 2), 98, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATO), 3, 12, 4, 2), 110, 12)
   Print #FISAL%, Z$
   Print #FISAL%, LINSEP$
   Close #FISAL%
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CALCOSOPE()
   '
   LOTESTA = Val(Text7.TEXT)
   If LOTESTA < 1 Then LOTESTA = 1
   CICLOHE = Val(Text34.TEXT)
   If CICLOHE < 1 Then CICLOHE = 1
   '
   CALOTES = 1
   MINPREP = 60 * Val(Text15(0)) + Val(Text17(0)) + Val(Text26(0)) / 60
   MINFIJO = 60 * Val(Text15(1)) + Val(Text17(1)) + Val(Text26(1)) / 60
   MINVARI = 60 * Val(Text15(2)) + Val(Text17(2)) + Val(Text26(2)) / 60
   MINHERR = LOTESTA * (60 * Val(Text15(4)) + Val(Text17(4)) + Val(Text26(4)) / 60) / CICLOHE
   MINDEMO = (MINFIJO + MINVARI + MINHERR) * Val(Text29.TEXT) / 100
   'MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
   MINLOTE = MINFIJO + MINVARI + MINHERR + MINDEMO
   HORAPROD = MINLOTE / 60 / LOTESTA
   '
   CAOPS = Val(Text30.TEXT)
   CATOP% = Val(Text33.TEXT)
   EQUIPO$ = Text9(0).TEXT
   CUOTOPE = VALHORA(CATOP%)
   CUOTMAQ = CUHORMAQ(EQUIPO$)
   '
   CIXI% = CODINT%(Text1): NOPYI% = Val(Label55)
   CAOPREP = CANOPREP(CIXI%, NOPYI%)
   If CAOPREP < 1 Then CAOPREP = CAOPS
   CTOPREP% = CATOPREP%(CIXI%, NOPYI%)
   If CTOPREP% < 1 Then CTOPREP% = CATOP%
   CUOTOPREP = VALHORA(CTOPREP%)
   '
   COSMOD = CUOTOPE * CAOPS * HORAPROD + CUOTOPREP * CAOPREP * MINPREP / 60 / LOTESTA
   COSMAQ = CUOTMAQ * (HORAPROD + MINPREP / 60 / LOTESTA)
   '
   Label73.Caption = TRIM$(CSTRING$(MKS$(COSMOD + COSMAQ), 3, 12, 4, 2))
   '
End Sub

Sub BLANFORCOS()
   '
   Label80 = ""
   Label85 = ""
   Label70 = ""
   Label90 = ""
   Label82 = ""
   Label88 = ""
   Label83 = ""
   Label86 = ""
   Label84 = ""
   Label75 = ""
   Label72 = ""
   Label77 = ""
   Label79 = ""
   Label74 = ""
   Label96 = ""
   Label103 = ""
   Label104 = ""
   Label105 = ""
   Label111 = ""
   Label112 = ""
   Label114 = ""
   PREVENETO# = 0
   Label116 = ""
   ANACOSMAT.List1.Clear
   ANACOSMOD.List1.Clear
   
   DoEvents
   '
End Sub

Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   Command26.SetFocus
End Sub
'
Sub COMULTI()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   CICON% = CI%
   '
   Screen.MousePointer = 11
   Call EXPLOMUL(CI%)
   If CONTAMUL% < 2 Then Exit Sub
   '
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   IMPLOSI.Show
   '
   COEFCA = 1: If PORMILLAR% = 1 Then COEFCA = 1000
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
     Call REPLA(Z$, SP$ + DESCRIT0$(CIK%(KKI%)), 17, 32)
     Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 51, 2)
     Call REPLA(Z$, CSTRING$(MKS$(COEFCA * CANT(KKI%)), 3, 10, 4, 2), 55, 10)
     IMPLOSI.List1.AddItem Z$
   Next KKI%
   '
   Screen.MousePointer = 1
   '
End Sub

Sub ACVEPROINT()
   For UI& = 0 To 32767
     PROVINT%(UI&) = 0
   Next UI&
   '
   For U& = 1 To ULTREG&("TAPROINT")
      X$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(X$, 2))
      If NP% > 0 Then
        PROVINT%(NP%) = 1
      End If
   Next U&
   Call CIERRARCH("TAPROINT")
   '
End Sub

Sub PRIFICHCOS()
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
   FORIPRE$ = "FANACOS"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        PRINTFORM
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Analisis de Costo " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        '
        CODPARAM$(5) = "USO-MPRI": DOCPARAM$(5) = Label90
        CODPARAM$(6) = "PIE-COMP": DOCPARAM$(6) = Label88
        CODPARAM$(7) = "DOMI-TRA": DOCPARAM$(7) = Label86
        CODPARAM$(8) = "CAN-MPRI": DOCPARAM$(8) = Label82
        CODPARAM$(9) = "CAN-COMP": DOCPARAM$(9) = Label83
        CODPARAM$(10) = "LOCA-TRA": DOCPARAM$(10) = Label84
        CODPARAM$(11) = "CANT-RES": DOCPARAM$(11) = Label85
        '
        CODPARAM$(12) = "TOT-PREP": DOCPARAM$(12) = Label75
        CODPARAM$(13) = "TOT-PROD": DOCPARAM$(13) = Label79
        CODPARAM$(14) = "TOT-DEBE": DOCPARAM$(14) = Label72
        CODPARAM$(15) = "TOT-HABE": DOCPARAM$(15) = Label74
        CODPARAM$(16) = "TOT-TIEM": DOCPARAM$(16) = Label80
        '
        CODPARAM$(17) = "PRE-LIST": DOCPARAM$(17) = Label114
        CODPARAM$(18) = "POR-DESC": DOCPARAM$(18) = Text44
        CODPARAM$(19) = "IMP-DESC": DOCPARAM$(19) = Label115
        CODPARAM$(20) = "PRE-NETO": DOCPARAM$(20) = Label112
        CODPARAM$(21) = "IMPNEDES": DOCPARAM$(21) = Label116
        CODPARAM$(22) = "NOP-OPER": DOCPARAM$(22) = Label117
        CODPARAM$(23) = "VALO-DEC": DOCPARAM$(23) = Label118
        CODPARAM$(24) = "DES-OPER": DOCPARAM$(24) = Label119
        CAPARDOC% = 24
        '
        CODTABU1$(1) = "POR-IVA"
        CODTABU1$(2) = "CANTIDAD"
        CODTABU1$(3) = "NET-ITEM"
        CACOLTAB1% = 3
        '
        For KK1% = 1 To 3
          For KK2% = 1 To 11
            TETABU1$(KK1%, KK2%) = ""
          Next KK2%
        Next KK1%
        '
        TETABU1$(3, 1) = Label70
        '
        TETABU1$(1, 2) = Text41
        TETABU1$(2, 2) = Label93
        '
        TETABU1$(1, 3) = Text36
        TETABU1$(2, 3) = Label68
        '
        TETABU1$(3, 4) = Label65
        '
        TETABU1$(1, 5) = Text38
        TETABU1$(2, 5) = Label77
        '
        TETABU1$(3, 6) = Label96
        '
        TETABU1$(1, 7) = Text42
        TETABU1$(2, 7) = Label110
        '
        TETABU1$(1, 8) = Text43
        TETABU1$(2, 8) = Label104
        '
        TETABU1$(3, 9) = Label111
        '
        TETABU1$(1, 10) = Text39
        TETABU1$(2, 10) = Label107
        '
        TETABU1$(3, 11) = Label103
        '
        CAITAB1% = 11
        '
        Call PRINTDOC("@FANACOS")
        Screen.MousePointer = 1
        '
      End If
      '
End Sub

Sub PRIANASUMI()
   '
   FORIPRE$ = "FANACOP"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        Exit Sub
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Costo Suministros " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = ANACOSMAT.Label3.Caption
        CODPARAM$(6) = "TOT-DEBE": DOCPARAM$(6) = ANACOSMAT.Label1.Caption
        CODPARAM$(7) = "TOT-HABE": DOCPARAM$(7) = ANACOSMAT.Label4.Caption
        CODPARAM$(8) = "TOT-PROD": DOCPARAM$(8) = ANACOSMAT.Label5.Caption
        CAPARDOC% = 8
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "DES-OPER"
        CODTABU1$(4) = "NUOPERS"
        CODTABU1$(5) = "TIE-PREP"
        CODTABU1$(6) = "PRE-LIST"
        CODTABU1$(7) = "IMP-DEBE"
        CODTABU1$(8) = "IMP-HABE"
        CODTABU1$(9) = "TIE-PROD"
        CODTABU1$(10) = "TIE-TOTA"
        CODTABU1$(11) = "IMP-TOTA"
        CACOLTAB1% = 11
        '
        CAITAB1% = 0: SUMAPAR# = 0
        For UI& = 1 To ANACOSMAT.List1.ListCount
          '
          TTXX$ = ANACOSMAT.List1.List(UI& - 1)
          If Left$(TTXX$, 64) = String$(64, "-") Then GoTo 19
          If Left$(TTXX$, 64) = String$(64, "=") Then GoTo 19
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Left$(TTXX$, 15)
          TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 129, 40)
          TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 40, 10)
          TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 52, 4)
          TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 169, 10)
          TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 179, 10)
          TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 189, 10)
          TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 55, 10)
          TETABU1$(9, CAITAB1%) = Mid$(TTXX$, 65, 10)
          TETABU1$(10, CAITAB1%) = Mid$(TTXX$, 75, 10)
          TETABU1$(11, CAITAB1%) = Mid$(TTXX$, 85, 10)
19      Next UI&
        '
        Call PRINTDOC("@FANACOP")
        Screen.MousePointer = 1
        '
      End If
      '
End Sub
   
Sub PRIANAMODI()
   '
   FORIPRE$ = "FANACOM"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        Exit Sub
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Costo M.O.D." + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = ANACOSMOD.Label3.Caption
        CODPARAM$(6) = "TOT-DEBE": DOCPARAM$(6) = ANACOSMOD.Label1.Caption
        CODPARAM$(7) = "TOT-HABE": DOCPARAM$(7) = ANACOSMOD.Label4.Caption
        CODPARAM$(8) = "TOT-PROD": DOCPARAM$(8) = ANACOSMOD.Label5.Caption
        CAPARDOC% = 8
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "DES-OPER"
        CODTABU1$(4) = "NUOPERS"
        CODTABU1$(5) = "TIE-PREP"
        CODTABU1$(6) = "PRE-LIST"
        CODTABU1$(7) = "IMP-DEBE"
        CODTABU1$(8) = "IMP-HABE"
        CODTABU1$(9) = "TIE-PROD"
        CODTABU1$(10) = "TIE-TOTA"
        CACOLTAB1% = 10
        '
        CAITAB1% = 0: SUMAPAR# = 0
        For UI& = 1 To ANACOSMOD.List1.ListCount
          '
          TTXX$ = ANACOSMOD.List1.List(UI& - 1)
          If Left$(TTXX$, 64) = String$(64, "-") Then GoTo 19
          If Left$(TTXX$, 64) = String$(64, "=") Then GoTo 19
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Left$(TTXX$, 15)
          TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 129, 40)
          TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 169, 30)
          TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 199, 5)
          TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 51, 14)
          TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 204, 10)
          TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 65, 10)
          TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 75, 10)
          TETABU1$(9, CAITAB1%) = Mid$(TTXX$, 85, 10)
          SUMAPAR# = SUMAPAR# + Val(Mid$(TTXX$, 85, 10))
          '
          TETABU1$(10, CAITAB1%) = ""
          TTYY$ = String$(64, "-")
          If UI& < ANACOSMOD.List1.ListCount Then
            TTYY$ = Left$(ANACOSMOD.List1.List(UI&), 64)
          End If
          If TTYY$ = String$(64, "-") Or TTYY$ = String$(64, "=") Then
            TTZZ$ = CSTRING$(MKD$(SUMAPAR#), 4, 10, DECICALCOS%, 2)
            TETABU1$(10, CAITAB1%) = TTZZ$
            SUMAPAR# = 0
            '
            CAITAB1% = CAITAB1% + 1
            For KKI% = 1 To 10
              TETABU1$(KKI%, CAITAB1%) = ""
            Next KKI%
          End If
          '
19      Next UI&
        '
        Call PRINTDOC("@FANACOM")
        Screen.MousePointer = 1
        '
      End If
      '

End Sub

Sub CARFORMULA()
   '
   If YACAR%(2) < 1 Then
      '
'Call ABRESTRUC
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      Call LEEEXPLO(CODD$, 1)
      PORMILLAR% = 0
      Call LEEEXPLO(CODD$, 1)
      Label59.Caption = FECHATEX$(REVIFORMU)
      List1.Clear
      NOCHECK% = 1
      Check1.Value = 0: If PORMILLAR% = 1 Then Check1.Value = 1
      NOCHECK% = 0
      '
      For I% = 1 To CAIT%
         List1.AddItem EXPLOLIN$(I%)
      Next I%
      INDEXI% = 0
      If IUCOD% >= 0 Then
        If ICOMPO%(IUCOD% + 1) >= 0 Then
          If ICOMPO%(IUCOD% + 1) < List1.ListCount Then
            INDEXI% = ICOMPO%(IUCOD% + 1)
          End If
        End If
      End If
      If List1.ListCount > 0 Then
         List1.ListIndex = INDEXI%
         If List1.TopIndex < List1.ListIndex - 8 Then
           List1.TopIndex = List1.ListIndex - 8
         End If
      End If
      '
      YACAR%(2) = 1
      MODIFIC%(2) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub


