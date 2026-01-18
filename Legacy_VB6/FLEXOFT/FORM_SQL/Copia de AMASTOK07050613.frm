VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form AMASTOK07 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos - A-B-M-L Maestro de Items de Stock"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11910
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   265
      Left            =   7440
      TabIndex        =   296
      Top             =   8040
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Frame Frame16 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   265
      Left            =   120
      TabIndex        =   295
      Top             =   8040
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.TextBox Text38 
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
      Left            =   1275
      TabIndex        =   274
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   256
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   3780
      Picture         =   "AMASTOK07.frx":0000
      TabIndex        =   4
      Top             =   225
      Width           =   175
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
      Height          =   600
      Left            =   1275
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   5295
   End
   Begin VB.TextBox Text45 
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
      Left            =   4725
      MaxLength       =   25
      TabIndex        =   127
      Top             =   210
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   165
      ToolTipText     =   "Tabla de Marcas y Modelos"
      Top             =   220
      Visible         =   0   'False
      Width           =   260
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Acc.Directos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Index           =   1
      Left            =   7665
      TabIndex        =   122
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
         Picture         =   "AMASTOK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   124
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
         Picture         =   "AMASTOK07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   8310
      Left            =   11025
      ScaleHeight     =   8250
      ScaleWidth      =   1005
      TabIndex        =   112
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command58 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   250
         ToolTipText     =   "Refresca Descripciones"
         Top             =   1480
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command61 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AMASTOK07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Actualizar Precios de Venta"
         Top             =   4200
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   168
         Top             =   7140
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   169
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMASTOK07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AMASTOK07.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AMASTOK07.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Subst"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AMASTOK07.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "AMASTOK07.frx":1150
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   3600
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
         Picture         =   "AMASTOK07.frx":145A
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
         Width           =   650
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AMASTOK07.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Imprime Ficha del Articulo"
         Top             =   6330
         Width           =   650
      End
      Begin VB.CommandButton Command38 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AMASTOK07.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   5640
         Width           =   650
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AMASTOK07.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4950
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMASTOK07.frx":251A
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AMASTOK07.frx":2824
         Top             =   7455
         Width           =   1515
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   5
      Top             =   240
      Width           =   2565
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   770
         Picture         =   "AMASTOK07.frx":4746
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
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
         Left            =   140
         Picture         =   "AMASTOK07.frx":4A50
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   630
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   13693168
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Insumos"
      TabPicture(0)   =   "AMASTOK07.frx":4D5A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame22"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MARCOFICHA(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Componentes"
      TabPicture(1)   =   "AMASTOK07.frx":4D76
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Elaboración"
      TabPicture(2)   =   "AMASTOK07.frx":4D92
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Especificación"
      TabPicture(3)   =   "AMASTOK07.frx":4DAE
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "MARCOFICHA(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame18"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame18 
         Caption         =   "Memoria Técnica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   70
         TabIndex        =   195
         Top             =   5520
         Width           =   10505
         Begin VB.CommandButton Command54 
            Caption         =   "X"
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
            Left            =   10110
            TabIndex        =   251
            ToolTipText     =   "Blanquear Casillero Memoria Técnica"
            Top             =   345
            Width           =   300
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   8820
            Top             =   210
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command54 
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
            Left            =   9840
            Picture         =   "AMASTOK07.frx":4DCA
            TabIndex        =   197
            Top             =   345
            Width           =   175
         End
         Begin VB.Label Label63 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   196
            Top             =   360
            Width           =   9630
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Height          =   6150
         Index           =   1
         Left            =   20
         TabIndex        =   172
         Top             =   315
         Width           =   10650
         Begin VB.TextBox Text43 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4035
            Left            =   70
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   298
            Top             =   720
            Width           =   7785
         End
         Begin VB.Frame FRAME17 
            Caption         =   "P.O.E.'S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1830
            Index           =   7
            Left            =   8040
            TabIndex        =   252
            Top             =   3255
            Width           =   2535
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   870
               Left            =   105
               TabIndex        =   255
               Top             =   210
               Width           =   2325
            End
            Begin VB.CommandButton Command47 
               Caption         =   "Quitar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1365
               TabIndex        =   254
               Top             =   1365
               Width           =   1125
            End
            Begin VB.CommandButton Command47 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   105
               TabIndex        =   253
               Top             =   1365
               Width           =   1185
            End
         End
         Begin VB.Frame FRAME17 
            Caption         =   "Batch-Card"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Index           =   3
            Left            =   8040
            TabIndex        =   214
            Top             =   1155
            Width           =   2535
            Begin VB.CommandButton Command69 
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
               Left            =   2265
               TabIndex        =   219
               Top             =   1560
               Width           =   175
            End
            Begin VB.CommandButton Command68 
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
               Left            =   2265
               TabIndex        =   217
               Top             =   960
               Width           =   175
            End
            Begin VB.TextBox Text84 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1320
               TabIndex        =   220
               Top             =   1575
               Width           =   975
            End
            Begin VB.TextBox Text83 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1320
               TabIndex        =   218
               Top             =   975
               Width           =   975
            End
            Begin VB.TextBox Text82 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   216
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha    Revisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   46
               Left            =   285
               TabIndex        =   224
               Top             =   1500
               Width           =   915
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha     Vigencia: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Index           =   45
               Left            =   285
               TabIndex        =   223
               Top             =   900
               Width           =   900
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Version:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   42
               Left            =   -360
               TabIndex        =   222
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.Frame FRAME17 
            Caption         =   "Clasificación ATC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   2
            Left            =   8040
            TabIndex        =   213
            Top             =   240
            Width           =   2535
            Begin VB.TextBox Text81 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   215
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Cortar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   0
            Left            =   2940
            TabIndex        =   179
            Top             =   4770
            Width           =   1170
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Copiar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   1
            Left            =   4095
            TabIndex        =   178
            Top             =   4770
            Width           =   1170
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Pegar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   2
            Left            =   5265
            TabIndex        =   177
            Top             =   4770
            Width           =   1275
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Selecc Todo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   3
            Left            =   6555
            TabIndex        =   176
            Top             =   4770
            Width           =   1320
         End
         Begin VB.CommandButton Command46 
            Height          =   260
            Left            =   75
            TabIndex        =   175
            Top             =   4770
            Width           =   2825
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   75
            ScaleHeight     =   375
            ScaleWidth      =   7725
            TabIndex        =   173
            Top             =   315
            Width           =   7780
            Begin VB.Label Label44 
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
               Height          =   225
               Left            =   0
               TabIndex        =   174
               Top             =   105
               Width           =   7680
            End
         End
         Begin VB.Line Line3 
            X1              =   7800
            X2              =   7800
            Y1              =   315
            Y2              =   5010
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Height          =   6150
         Index           =   3
         Left            =   -74980
         TabIndex        =   52
         Top             =   310
         Width           =   10680
         Begin VB.Frame Frame4 
            Caption         =   "Detalle de la Operación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5925
            Left            =   4850
            TabIndex        =   56
            Top             =   165
            Width           =   5685
            Begin VB.Frame FRAME17 
               Caption         =   "RRHH - Personal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   950
               Index           =   6
               Left            =   0
               TabIndex        =   261
               Top             =   5000
               Width           =   5685
               Begin VB.TextBox Text30 
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
                  Height          =   285
                  Index           =   0
                  Left            =   1215
                  TabIndex        =   267
                  Top             =   240
                  Width           =   570
               End
               Begin VB.TextBox Text33 
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
                  Height          =   285
                  Index           =   0
                  Left            =   2370
                  TabIndex        =   266
                  Top             =   240
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
                  Index           =   0
                  Left            =   2730
                  TabIndex        =   265
                  Top             =   240
                  Width           =   175
               End
               Begin VB.TextBox Text30 
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
                  Height          =   285
                  Index           =   1
                  Left            =   1215
                  TabIndex        =   264
                  Top             =   550
                  Width           =   570
               End
               Begin VB.TextBox Text33 
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
                  Height          =   285
                  Index           =   1
                  Left            =   2370
                  TabIndex        =   263
                  Top             =   550
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
                  Index           =   1
                  Left            =   2730
                  TabIndex        =   262
                  Top             =   550
                  Width           =   175
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cat:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   61
                  Left            =   1965
                  TabIndex        =   273
                  Top             =   600
                  Width           =   330
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Operarios:              Cat:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   15
                  Left            =   240
                  TabIndex        =   272
                  Top             =   300
                  Width           =   855
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
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   271
                  Top             =   240
                  Width           =   2460
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Operarios:      "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   44
                  Left            =   240
                  TabIndex        =   270
                  Top             =   600
                  Width           =   960
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
                  Index           =   1
                  Left            =   3000
                  TabIndex        =   269
                  Top             =   550
                  Width           =   2460
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cat:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   60
                  Left            =   1965
                  TabIndex        =   268
                  Top             =   300
                  Width           =   330
               End
            End
            Begin VB.CommandButton Command70 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4920
               Picture         =   "AMASTOK07.frx":50D4
               Style           =   1  'Graphical
               TabIndex        =   247
               ToolTipText     =   "Herramentales/Equipos Auxiliares"
               Top             =   4440
               Width           =   615
            End
            Begin VB.Frame FRAME17 
               Caption         =   "Equipos Auxiliares"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2110
               Index           =   5
               Left            =   3290
               TabIndex        =   231
               Top             =   2280
               Width           =   2235
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
                  Height          =   285
                  Index           =   5
                  Left            =   1935
                  TabIndex        =   244
                  Top             =   1600
                  Width           =   175
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
                  Height          =   285
                  Index           =   4
                  Left            =   1935
                  TabIndex        =   243
                  Top             =   1290
                  Width           =   175
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
                  Index           =   5
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   242
                  Top             =   1600
                  Width           =   780
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
                  Index           =   4
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   241
                  Top             =   1290
                  Width           =   780
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
                  Height          =   285
                  Index           =   3
                  Left            =   1935
                  TabIndex        =   237
                  Top             =   980
                  Width           =   175
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
                  Height          =   285
                  Index           =   2
                  Left            =   1935
                  TabIndex        =   236
                  Top             =   670
                  Width           =   175
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
                  Height          =   285
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   235
                  Top             =   360
                  Width           =   175
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
                  Index           =   3
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   234
                  Top             =   980
                  Width           =   780
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
                  Index           =   2
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   233
                  Top             =   670
                  Width           =   780
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
                  Index           =   1
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   232
                  Top             =   360
                  Width           =   780
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eq.Aux (5):"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   5
                  Left            =   120
                  TabIndex        =   246
                  Top             =   1660
                  Width           =   960
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eq.Aux (4):"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   4
                  Left            =   120
                  TabIndex        =   245
                  Top             =   1330
                  Width           =   960
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eq.Aux (3):"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   3
                  Left            =   120
                  TabIndex        =   240
                  Top             =   1040
                  Width           =   960
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eq.Aux (2):"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   2
                  Left            =   120
                  TabIndex        =   239
                  Top             =   730
                  Width           =   960
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eq.Aux (1):"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   1
                  Left            =   120
                  TabIndex        =   238
                  Top             =   420
                  Width           =   960
               End
            End
            Begin VB.Frame FRAME17 
               Caption         =   "% D-D"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   680
               Index           =   4
               Left            =   1800
               TabIndex        =   229
               Top             =   3720
               Width           =   1275
               Begin VB.TextBox Text29 
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
                  Height          =   285
                  Left            =   360
                  TabIndex        =   230
                  ToolTipText     =   "Porcentaje Demora Distribuida"
                  Top             =   285
                  Width           =   570
               End
            End
            Begin VB.Frame FRAME17 
               Caption         =   "Cadencia"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   680
               Index           =   0
               Left            =   210
               TabIndex        =   183
               Top             =   3720
               Width           =   1515
               Begin VB.TextBox Text47 
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
                  Height          =   285
                  Left            =   360
                  TabIndex        =   185
                  ToolTipText     =   "Piezas / Hora"
                  Top             =   280
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label Label50 
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
                  Left            =   360
                  TabIndex        =   184
                  ToolTipText     =   "Doble-Click para Ingresar Cadencia"
                  Top             =   280
                  Width           =   885
               End
            End
            Begin VB.TextBox Text34 
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
               Height          =   285
               Left            =   4200
               TabIndex        =   164
               Text            =   " "
               Top             =   4560
               Width           =   675
            End
            Begin VB.TextBox Text15 
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
               Height          =   285
               Index           =   3
               Left            =   1260
               TabIndex        =   162
               Top             =   3330
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Height          =   285
               Index           =   3
               Left            =   1785
               TabIndex        =   161
               Top             =   3330
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Height          =   285
               Index           =   3
               Left            =   2205
               TabIndex        =   160
               Top             =   3330
               Width           =   930
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
               Left            =   3675
               TabIndex        =   151
               Top             =   1440
               Width           =   1830
            End
            Begin VB.TextBox Text37 
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
               Left            =   4200
               TabIndex        =   149
               Text            =   " "
               Top             =   4560
               Width           =   675
            End
            Begin VB.CommandButton Command28 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   2100
               TabIndex        =   60
               Top             =   1440
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
               TabIndex        =   58
               Top             =   720
               Width           =   4245
            End
            Begin VB.TextBox Text7 
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
               Height          =   285
               Left            =   4410
               TabIndex        =   57
               Top             =   405
               Width           =   1095
            End
            Begin VB.TextBox Text9 
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
               Height          =   285
               Index           =   0
               Left            =   1260
               MaxLength       =   6
               TabIndex        =   59
               Top             =   1440
               Width           =   885
            End
            Begin VB.TextBox Text15 
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
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   61
               Top             =   2385
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Height          =   285
               Index           =   0
               Left            =   1785
               TabIndex        =   62
               Top             =   2385
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Height          =   285
               Index           =   0
               Left            =   2205
               TabIndex        =   63
               Top             =   2385
               Width           =   930
            End
            Begin VB.TextBox Text15 
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
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   64
               Top             =   2700
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Height          =   285
               Index           =   1
               Left            =   1785
               TabIndex        =   65
               Top             =   2700
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Height          =   285
               Index           =   1
               Left            =   2205
               TabIndex        =   66
               Top             =   2700
               Width           =   930
            End
            Begin VB.TextBox Text15 
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
               Height          =   285
               Index           =   2
               Left            =   1260
               TabIndex        =   67
               Top             =   3015
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Height          =   285
               Index           =   2
               Left            =   1785
               TabIndex        =   68
               Top             =   3015
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Height          =   285
               Index           =   2
               Left            =   2205
               TabIndex        =   69
               Top             =   3015
               Width           =   930
            End
            Begin VB.TextBox Text15 
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
               Height          =   285
               Index           =   4
               Left            =   1260
               TabIndex        =   70
               Top             =   4560
               Width           =   465
            End
            Begin VB.TextBox Text17 
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
               Height          =   285
               Index           =   4
               Left            =   1785
               TabIndex        =   71
               Top             =   4560
               Width           =   360
            End
            Begin VB.TextBox Text26 
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
               Height          =   285
               Index           =   4
               Left            =   2205
               TabIndex        =   72
               Top             =   4560
               Width           =   930
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Standard:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   59
               Left            =   2800
               TabIndex        =   260
               Top             =   480
               Width           =   1275
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Resp:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   58
               Left            =   2475
               TabIndex        =   259
               Top             =   1110
               Width           =   495
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "Ciclo M-H:         "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   3255
               TabIndex        =   257
               ToolTipText     =   "Doble Click para Detalle Mant.Herramental"
               Top             =   4590
               Width           =   945
            End
            Begin VB.Label Label52 
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
               Left            =   1935
               TabIndex        =   187
               Top             =   405
               Width           =   300
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "HPF:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   2835
               TabIndex        =   163
               Top             =   1545
               Width           =   750
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "HPF:"
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
               Left            =   2835
               TabIndex        =   152
               Top             =   11785
               Width           =   750
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "T.Limpieza:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   3
               Left            =   105
               TabIndex        =   150
               Top             =   3425
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00D8BD05&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4200
               TabIndex        =   129
               ToolTipText     =   "Tiempos por Unidad"
               Top             =   405
               Width           =   165
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   420
               Picture         =   "AMASTOK07.frx":53DE
               ToolTipText     =   "Abrir Tiempos de Preparación"
               Top             =   2070
               Width           =   480
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Operación:           -                   Lote Standard:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   13
               Left            =   255
               TabIndex        =   90
               Top             =   500
               Width           =   960
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
               Left            =   1260
               TabIndex        =   89
               Top             =   1755
               Width           =   4245
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo:"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   88
               Top             =   1564
               Width           =   960
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Preparación:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   105
               TabIndex        =   87
               ToolTipText     =   "Doble-Click para Abrir Tiempos de Preparación"
               Top             =   2480
               Width           =   1095
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " T.Std.Fijo:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   105
               TabIndex        =   86
               Top             =   2795
               Width           =   1095
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "T.Var.Lote:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   105
               TabIndex        =   85
               ToolTipText     =   "Doble-Click para Abrir Tiempos Variables por Lote"
               Top             =   3110
               Width           =   1095
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "       H.    M.    Seg."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   200
               Index           =   43
               Left            =   1260
               TabIndex        =   83
               Top             =   2160
               Width           =   1800
            End
            Begin VB.Label Label55 
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
               Left            =   1260
               TabIndex        =   82
               Top             =   405
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
               TabIndex        =   81
               Top             =   1035
               Width           =   885
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:                      Resp:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   14
               Left            =   375
               TabIndex        =   80
               Top             =   1130
               Width           =   810
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
               TabIndex        =   79
               Top             =   1035
               Width           =   2460
            End
            Begin VB.Label Label47 
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Mant.Herr:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   84
               ToolTipText     =   "Doble Click para Detalle Mant.Herramental"
               Top             =   4560
               Width           =   1125
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Operaciones Standard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Left            =   160
            TabIndex        =   54
            Top             =   165
            Width           =   4530
            Begin VB.CommandButton Command53 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Alt"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   1755
               Style           =   1  'Graphical
               TabIndex        =   186
               ToolTipText     =   "Inserta Operación Alternativa a la Activa"
               Top             =   2520
               Width           =   550
            End
            Begin VB.CommandButton Command32 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   1205
               Style           =   1  'Graphical
               TabIndex        =   75
               ToolTipText     =   "Elimina la Operación Activa"
               Top             =   2520
               Width           =   550
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ins"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   655
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
               Top             =   2520
               Width           =   550
            End
            Begin VB.CommandButton Command551 
               Caption         =   "COPY"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3015
               TabIndex        =   182
               ToolTipText     =   "Copia de TMU's"
               Top             =   180
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.CommandButton Command555 
               Caption         =   "TMU's"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3725
               TabIndex        =   181
               ToolTipText     =   "Listado de TMU's"
               Top             =   180
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   73
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   2520
               Width           =   550
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
               Height          =   1860
               Left            =   105
               TabIndex        =   55
               Top             =   420
               Width           =   4320
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   2520
               Width           =   845
            End
            Begin VB.CommandButton Command42 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Listar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   3580
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   2520
               Width           =   845
            End
         End
         Begin VB.Frame FRAME17 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Index           =   1
            Left            =   160
            TabIndex        =   53
            Top             =   3315
            Width           =   4530
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2340
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               Top             =   315
               Width           =   4320
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6160
         Index           =   2
         Left            =   -74980
         TabIndex        =   41
         Top             =   310
         Width           =   10680
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2160
            Left            =   390
            TabIndex        =   276
            Top             =   3920
            Width           =   8100
            Begin VB.Frame Frame3 
               Caption         =   "Datos Componente Activo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2160
               Left            =   0
               TabIndex        =   299
               Top             =   0
               Width           =   7275
               Begin VB.CommandButton Command45 
                  Caption         =   "X"
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
                  Left            =   7035
                  TabIndex        =   304
                  Top             =   140
                  Width           =   225
               End
               Begin VB.CommandButton Command17 
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
                  Height          =   270
                  Left            =   3360
                  TabIndex        =   303
                  Top             =   475
                  Width           =   175
               End
               Begin VB.CheckBox Check2 
                  Caption         =   "x Millar"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   210
                  TabIndex        =   302
                  Top             =   1785
                  Width           =   960
               End
               Begin VB.TextBox Text42 
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
                  Height          =   255
                  Left            =   4740
                  TabIndex        =   301
                  Text            =   " "
                  Top             =   840
                  Width           =   750
               End
               Begin VB.TextBox Text41 
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
                  Height          =   255
                  Left            =   1245
                  TabIndex        =   300
                  Text            =   " "
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "U-M: "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   52
                  Left            =   3570
                  TabIndex        =   326
                  Top             =   525
                  Width           =   480
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "D:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   54
                  Left            =   6510
                  TabIndex        =   325
                  Top             =   525
                  Width           =   270
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Uso:                       D:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   53
                  Left            =   4830
                  TabIndex        =   324
                  Top             =   525
                  Width           =   375
               End
               Begin VB.Label Label23 
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
                  Height          =   540
                  Left            =   1240
                  TabIndex        =   323
                  Top             =   1455
                  Width           =   5915
               End
               Begin VB.Label Label22 
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
                  Height          =   270
                  Left            =   6825
                  TabIndex        =   322
                  Top             =   480
                  Width           =   330
               End
               Begin VB.Label Label20 
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
                  Height          =   270
                  Left            =   5250
                  TabIndex        =   321
                  Top             =   480
                  Width           =   1170
               End
               Begin VB.Label Label19 
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
                  Height          =   270
                  Left            =   1260
                  TabIndex        =   320
                  Top             =   480
                  Width           =   2050
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código:                                          U-M:           Uso:                       D:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   18
                  Left            =   540
                  TabIndex        =   319
                  Top             =   525
                  Width           =   690
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Descripción:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   17
                  Left            =   105
                  TabIndex        =   318
                  Top             =   840
                  Width           =   1065
               End
               Begin VB.Label Label9 
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
                  Height          =   270
                  Left            =   4200
                  TabIndex        =   317
                  Top             =   480
                  Width           =   435
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1240
                  TabIndex        =   316
                  Top             =   840
                  Width           =   5905
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Referencia:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   16
                  Left            =   105
                  TabIndex        =   315
                  Top             =   1450
                  Width           =   1065
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1245
                  TabIndex        =   314
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cantidad:"
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
                  Left            =   120
                  TabIndex        =   313
                  Top             =   885
                  Width           =   1065
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3090
                  TabIndex        =   312
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kg/M:"
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
                  Left            =   2400
                  TabIndex        =   311
                  Top             =   885
                  Width           =   645
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4740
                  TabIndex        =   310
                  Top             =   840
                  Width           =   750
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "L(mm):"
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
                  Left            =   4050
                  TabIndex        =   309
                  Top             =   885
                  Width           =   645
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6210
                  TabIndex        =   308
                  Top             =   840
                  Width           =   930
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kg/U:"
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
                  Left            =   5625
                  TabIndex        =   307
                  Top             =   885
                  Width           =   570
               End
               Begin VB.Label Label36 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6210
                  TabIndex        =   306
                  Top             =   840
                  Width           =   930
               End
               Begin VB.Label Label42 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3090
                  TabIndex        =   305
                  Top             =   840
                  Width           =   855
               End
            End
            Begin VB.CheckBox Check1 
               Caption         =   "x Millar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   210
               TabIndex        =   286
               Top             =   1785
               Width           =   960
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7400
               Picture         =   "AMASTOK07.frx":56E8
               TabIndex        =   285
               ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
               Top             =   1650
               Width           =   550
            End
            Begin VB.CommandButton Command31 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3360
               TabIndex        =   284
               Top             =   475
               Width           =   175
            End
            Begin VB.CommandButton Command19 
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7400
               Picture         =   "AMASTOK07.frx":59F2
               TabIndex        =   283
               ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
               Top             =   1260
               Width           =   550
            End
            Begin VB.TextBox Text27 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   6825
               TabIndex        =   282
               Text            =   " "
               Top             =   475
               Width           =   330
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Ins"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7400
               Picture         =   "AMASTOK07.frx":5CFC
               TabIndex        =   281
               ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
               Top             =   855
               Width           =   550
            End
            Begin VB.TextBox Text31 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1260
               TabIndex        =   280
               Top             =   475
               Width           =   2115
            End
            Begin VB.TextBox Text32 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5250
               TabIndex        =   279
               Text            =   " "
               Top             =   475
               Width           =   1170
            End
            Begin VB.CommandButton Command27 
               Caption         =   "App"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7400
               Picture         =   "AMASTOK07.frx":6006
               TabIndex        =   278
               ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
               Top             =   470
               Width           =   550
            End
            Begin VB.TextBox Text39 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   1240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   277
               Top             =   1450
               Width           =   5930
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Código:                                          U-M:           Uso:                       D:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   540
               TabIndex        =   294
               Top             =   525
               Width           =   690
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción:"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   293
               Top             =   840
               Width           =   1065
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4200
               TabIndex        =   292
               Top             =   480
               Width           =   435
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   1240
               TabIndex        =   291
               Top             =   840
               Width           =   5915
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Referencia:"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   290
               Top             =   1450
               Width           =   1065
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "U-M:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   55
               Left            =   3570
               TabIndex        =   289
               Top             =   525
               Width           =   480
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Uso:                       D:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   56
               Left            =   4830
               TabIndex        =   288
               Top             =   525
               Width           =   375
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "D:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   57
               Left            =   6510
               TabIndex        =   287
               Top             =   525
               Width           =   270
            End
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Sort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9705
            TabIndex        =   191
            ToolTipText     =   "Ordenamiento de Estructura"
            Top             =   405
            Width           =   525
         End
         Begin VB.Frame Frame12 
            Caption         =   "Consultas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1210
            Left            =   8685
            TabIndex        =   107
            Top             =   4860
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Implosión"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   109
               ToolTipText     =   "Consulta de Conjuntos a los que Pertenece el Elemento Activo"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   108
               ToolTipText     =   "Muestra Estructura Multinivel"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   390
            ScaleHeight     =   375
            ScaleWidth      =   8040
            TabIndex        =   51
            Top             =   420
            Width           =   8100
            Begin VB.Label Label27 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   375
               Left            =   0
               TabIndex        =   258
               Top             =   0
               Width           =   8040
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
            Left            =   390
            TabIndex        =   50
            Top             =   855
            Width           =   8100
         End
         Begin VB.Frame Frame11 
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   8685
            TabIndex        =   46
            Top             =   3315
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Enlaces"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   49
               ToolTipText     =   "Listado de Estucturas y Enlaces"
               Top             =   380
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   48
               ToolTipText     =   "Lista de Materiales Multinivel"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "L.Compras"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   47
               ToolTipText     =   "Lista de Compras"
               Top             =   1100
               Width           =   1170
            End
         End
         Begin VB.CommandButton Command41 
            Height          =   435
            Left            =   8685
            Picture         =   "AMASTOK07.frx":6310
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   405
            Width           =   1005
         End
         Begin VB.Frame Frame14 
            Caption         =   "Utilidades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   8685
            TabIndex        =   44
            Top             =   2145
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Control"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   106
               ToolTipText     =   "Control de Integridad de la Base de Datos"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   105
               ToolTipText     =   "Nueva Fórmula por Copia de una Pre-Existente"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame FRAME15 
            Caption         =   "Revisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   8685
            TabIndex        =   42
            Top             =   975
            Width           =   1590
            Begin VB.CommandButton Command51 
               Caption         =   "Aprobar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   210
               TabIndex        =   192
               Top             =   705
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label59 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
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
               Left            =   210
               TabIndex        =   43
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Line Line1 
            X1              =   8490
            X2              =   8490
            Y1              =   450
            Y2              =   1140
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Height          =   6150
         Index           =   0
         Left            =   -74980
         TabIndex        =   12
         Top             =   325
         Width           =   10660
         Begin VB.Frame Frame7 
            Caption         =   "Reposición"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2440
            Index           =   1
            Left            =   5280
            TabIndex        =   139
            Top             =   300
            Width           =   5160
            Begin VB.TextBox Text52 
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
               Height          =   330
               Left            =   2400
               MaxLength       =   4
               TabIndex        =   248
               ToolTipText     =   "Autonomía Proyectada de Cálculo"
               Top             =   2010
               Width           =   495
            End
            Begin VB.CommandButton Command56 
               Caption         =   "!"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1455
               TabIndex        =   228
               ToolTipText     =   "Actualizacion Masiva por Proveedor"
               Top             =   2010
               Width           =   180
            End
            Begin VB.TextBox Text51 
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
               Height          =   330
               Left            =   1080
               MaxLength       =   2
               TabIndex        =   227
               ToolTipText     =   "Cantidad dias Lead Time"
               Top             =   2010
               Width           =   375
            End
            Begin VB.TextBox Text49 
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
               Height          =   285
               Left            =   3600
               TabIndex        =   202
               Top             =   1350
               Width           =   1305
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Descomposición Automática"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   430
               Left            =   240
               TabIndex        =   180
               Top             =   1080
               Width           =   1905
            End
            Begin VB.TextBox Text40 
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
               Height          =   315
               Left            =   3600
               TabIndex        =   170
               Top             =   950
               Width           =   1305
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
               Height          =   285
               Left            =   3600
               TabIndex        =   158
               Top             =   240
               Width           =   1305
            End
            Begin VB.TextBox Text19 
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
               Height          =   285
               Left            =   3600
               TabIndex        =   148
               Top             =   600
               Width           =   1305
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
               Height          =   315
               Left            =   1995
               TabIndex        =   142
               Top             =   420
               Width           =   175
            End
            Begin VB.TextBox Text20 
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
               Height          =   285
               Left            =   3600
               TabIndex        =   141
               Top             =   1710
               Width           =   1305
            End
            Begin VB.TextBox Text8 
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
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   140
               Text            =   " "
               Top             =   420
               Width           =   360
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "AUT.:"
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
               Index           =   51
               Left            =   1720
               TabIndex        =   249
               ToolTipText     =   "Autonomía Proyectada de Cálculo"
               Top             =   2130
               Width           =   615
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "LTC:"
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
               Index           =   50
               Left            =   240
               TabIndex        =   226
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   2130
               Width           =   615
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Activo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   160
               Left            =   525
               TabIndex        =   221
               Top             =   1800
               Width           =   600
            End
            Begin VB.Label Label71 
               Alignment       =   2  'Center
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   212
               Top             =   1800
               Width           =   195
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Prom:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   47
               Left            =   3075
               TabIndex        =   207
               Top             =   2140
               Width           =   540
            End
            Begin VB.Label Label67 
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
               Left            =   3600
               TabIndex        =   206
               Top             =   2070
               Width           =   1305
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L. Indivis.Vta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   33
               Left            =   2160
               TabIndex        =   203
               Top             =   1450
               Width           =   1395
            End
            Begin VB.Label Label61 
               BackStyle       =   0  'Transparent
               Caption         =   "Trazable"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   525
               TabIndex        =   190
               Top             =   1500
               Width           =   1890
            End
            Begin VB.Label Label58 
               Alignment       =   2  'Center
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   200
               Left            =   240
               TabIndex        =   189
               Top             =   1500
               Width           =   200
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L. Indivis.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   34
               Left            =   2415
               TabIndex        =   171
               Top             =   1060
               Width           =   1155
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Máximo:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   32
               Left            =   2415
               TabIndex        =   147
               Top             =   1780
               Width           =   1155
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "St. Minimo:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   36
               Left            =   2415
               TabIndex        =   146
               Top             =   350
               Width           =   1155
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L. Económ.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   35
               Left            =   2415
               TabIndex        =   145
               Top             =   695
               Width           =   1155
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Criterio de Reposicion:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   37
               Left            =   -105
               TabIndex        =   144
               Top             =   320
               Width           =   1380
            End
            Begin VB.Label Label32 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   143
               Top             =   780
               Width           =   1920
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Costos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3140
            Index           =   2
            Left            =   5280
            TabIndex        =   21
            Top             =   2750
            Width           =   5160
            Begin VB.TextBox Text28 
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
               Height          =   285
               Left            =   3720
               TabIndex        =   111
               Top             =   1005
               Width           =   1150
            End
            Begin VB.TextBox Text21 
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
               Height          =   285
               Left            =   3465
               TabIndex        =   39
               Top             =   1890
               Width           =   1410
            End
            Begin VB.TextBox Text13 
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
               Height          =   285
               Left            =   1260
               MaxLength       =   2
               TabIndex        =   31
               Top             =   1005
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
               Top             =   1005
               Width           =   175
            End
            Begin VB.TextBox Text14 
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
               TabIndex        =   28
               Top             =   525
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
               Top             =   525
               Width           =   175
            End
            Begin VB.TextBox Text12 
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
               Height          =   285
               Left            =   1260
               TabIndex        =   25
               Top             =   525
               Width           =   1410
            End
            Begin VB.TextBox Text10 
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
               Height          =   285
               Left            =   1260
               MaxLength       =   5
               TabIndex        =   23
               Top             =   1890
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
               Top             =   1890
               Width           =   175
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   " % G.Nacion:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   62
               Left            =   2570
               TabIndex        =   275
               Top             =   1060
               Width           =   1120
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   31
               Left            =   2800
               TabIndex        =   211
               Top             =   2785
               Width           =   660
            End
            Begin VB.Label Label70 
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
               Left            =   3510
               TabIndex        =   210
               Top             =   2700
               Width           =   1410
            End
            Begin VB.Label Label69 
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
               Left            =   1260
               TabIndex        =   209
               Top             =   2685
               Width           =   1410
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pr. Básico (Costos):"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   30
               Left            =   160
               TabIndex        =   208
               ToolTipText     =   "l"
               Top             =   2560
               Width           =   1035
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "$ / Unidad:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   10
               Left            =   2310
               TabIndex        =   40
               Top             =   1965
               Width           =   1080
            End
            Begin VB.Label Label35 
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
               TabIndex        =   34
               Top             =   1320
               Width           =   3615
            End
            Begin VB.Label Label33 
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
               TabIndex        =   33
               Top             =   2265
               Width           =   3615
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   8
               Left            =   500
               TabIndex        =   32
               Top             =   1110
               Width           =   765
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Costeo:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   6
               Left            =   2715
               TabIndex        =   29
               Top             =   420
               Width           =   930
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Reposición:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   7
               Left            =   210
               TabIndex        =   26
               Top             =   420
               Width           =   1005
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedor Habitual:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   9
               Left            =   0
               TabIndex        =   24
               ToolTipText     =   "Doble-Click para Control Cambios Proveedor Habitual"
               Top             =   1785
               Width           =   1200
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Características"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5580
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   290
            Width           =   4740
            Begin VB.TextBox Text50 
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
               Height          =   330
               Left            =   4155
               MaxLength       =   1
               TabIndex        =   204
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   1335
               Width           =   375
            End
            Begin VB.TextBox Text48 
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
               Height          =   315
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   199
               Text            =   " "
               Top             =   840
               Width           =   570
            End
            Begin VB.CommandButton Command55 
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
               Left            =   1700
               TabIndex        =   198
               Top             =   840
               Width           =   175
            End
            Begin VB.CommandButton Command52 
               Caption         =   "Más Datos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3465
               TabIndex        =   194
               ToolTipText     =   "Actualización y Listado"
               Top             =   2850
               Width           =   1140
            End
            Begin VB.TextBox Text44 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1155
               TabIndex        =   155
               Top             =   2850
               Width           =   2250
            End
            Begin VB.TextBox Text46 
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
               Height          =   330
               Left            =   1155
               TabIndex        =   166
               Top             =   4455
               Width           =   1215
            End
            Begin VB.TextBox Text16 
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
               Height          =   315
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   159
               Text            =   " "
               Top             =   1820
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
               Height          =   300
               Left            =   1695
               TabIndex        =   156
               Top             =   1820
               Width           =   175
            End
            Begin VB.TextBox Text35 
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
               Height          =   315
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   154
               Text            =   "  "
               Top             =   2315
               Width           =   570
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
               Height          =   300
               Left            =   1695
               TabIndex        =   153
               Top             =   2315
               Width           =   175
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
               Height          =   300
               Left            =   4340
               TabIndex        =   136
               Top             =   5040
               Width           =   175
            End
            Begin VB.TextBox Text25 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3210
               TabIndex        =   135
               Text            =   " "
               Top             =   5040
               Width           =   1110
            End
            Begin VB.TextBox Text36 
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
               Left            =   1155
               TabIndex        =   134
               Top             =   5040
               Width           =   780
            End
            Begin VB.CommandButton Command44 
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
               Left            =   4410
               TabIndex        =   133
               Top             =   3360
               Width           =   175
            End
            Begin VB.CommandButton Command43 
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
               Left            =   4410
               TabIndex        =   132
               Top             =   3900
               Width           =   175
            End
            Begin VB.TextBox Text23 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1155
               TabIndex        =   102
               Top             =   3360
               Width           =   3240
            End
            Begin VB.TextBox Text22 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1155
               TabIndex        =   101
               Top             =   3900
               Width           =   3240
            End
            Begin VB.TextBox Text24 
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
               Left            =   3210
               TabIndex        =   20
               Top             =   4455
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
               Height          =   300
               Left            =   1700
               TabIndex        =   17
               Top             =   1320
               Width           =   175
            End
            Begin VB.TextBox Text4 
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
               Height          =   315
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   16
               Text            =   " "
               Top             =   1340
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
               Height          =   300
               Left            =   1700
               TabIndex        =   15
               Top             =   380
               Width           =   175
            End
            Begin VB.TextBox Text6 
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
               Height          =   315
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   14
               Text            =   " "
               Top             =   360
               Width           =   570
            End
            Begin VB.Label Label57 
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
               Height          =   315
               Left            =   1995
               TabIndex        =   225
               Top             =   2325
               Width           =   2535
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dec:"
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
               Left            =   3510
               TabIndex        =   205
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   1425
               Width           =   615
            End
            Begin VB.Label Label64 
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
               Height          =   315
               Left            =   1995
               TabIndex        =   201
               Top             =   840
               Width           =   2535
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo de Familias:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   29
               Left            =   -345
               TabIndex        =   200
               Top             =   760
               Width           =   1380
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cod.Intern.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   0
               TabIndex        =   193
               Top             =   2960
               Width           =   1095
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kilos/ Metro:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   19
               Left            =   195
               TabIndex        =   167
               Top             =   4400
               Width           =   870
            End
            Begin VB.Label Label31 
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
               Height          =   315
               Left            =   1995
               TabIndex        =   157
               Top             =   1820
               Width           =   2535
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Revisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   49
               Left            =   2205
               TabIndex        =   138
               Top             =   4970
               Width           =   960
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   48
               Left            =   -1575
               TabIndex        =   137
               Top             =   5170
               Width           =   2655
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Contable:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   -210
               TabIndex        =   131
               Top             =   1740
               Width           =   1290
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo I.V.A.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   2
               Left            =   120
               TabIndex        =   130
               Top             =   2220
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "        Color:"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   104
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Talle ó Medida:"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   103
               Top             =   3280
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Material:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   280
               Width           =   1005
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unidad de Medida:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   -105
               TabIndex        =   37
               Top             =   1270
               Width           =   1185
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kilos/ Unid:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   4
               Left            =   2550
               TabIndex        =   36
               Top             =   4400
               Width           =   630
            End
            Begin VB.Label Label30 
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
               Height          =   315
               Left            =   1995
               TabIndex        =   19
               Top             =   1335
               Width           =   1620
            End
            Begin VB.Label Label29 
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
               Height          =   315
               Left            =   1995
               TabIndex        =   18
               Top             =   380
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "No Usados"
         Height          =   4215
         Left            =   -74475
         TabIndex        =   91
         Top             =   630
         Width           =   4320
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
            TabIndex        =   95
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
            TabIndex        =   94
            Top             =   1710
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
            Top             =   0
            Width           =   1380
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   93
            Top             =   630
            Width           =   975
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
            TabIndex        =   92
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "AMASTOK07.frx":661A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "AMASTOK07.frx":684E
      TabIndex        =   297
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Pl.Nro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   41
      Left            =   4125
      TabIndex        =   128
      Top             =   285
      Width           =   555
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   40
      Left            =   6680
      TabIndex        =   126
      Top             =   675
      Width           =   520
   End
   Begin VB.Line Line8 
      X1              =   -120
      X2              =   11745
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   105
      TabIndex        =   110
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6900
      TabIndex        =   35
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
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
   Begin VB.Menu ConGen 
      Caption         =   "Configuración "
   End
End
Attribute VB_Name = "AMASTOK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%(15), MODIFIC%(15)
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
Dim PORLARGO%
'
Dim RESECO$(1024), DESOPE$(1024)
Dim MAXECO%
'
Dim PROVINT%(32767)
Dim FOCOREP%
'
Dim EXISTMU% '\\\OT-06-0252-WAR-26/06/06///
'
Dim PORPATRON%

Sub ACTUFORMUPAT(CI%, LOTEANTERIOR, LOTENUEVO)
     '
     CONDI$ = "CodIConj = " & CI%
     Call ACTUREGISTRO("FORMUPAT", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
     
     'VALIDO SI EXISTE EN LA FORMULA PATRON, SI NO SE VA.
     If RAFE& < 1 Then Exit Sub
     '
     SQLX$ = "SELECT  * FROM FORMULAS WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CodIConj = " & CI%
     Dim FORMUTMP As New ADODB.Recordset
     Set FORMUTMP = FLEXCONN.Execute(SQLX$)
     '
     With FORMUTMP
        Do While Not .EOF
           USO# = XVALO(!USOBRUTO)
           USO# = USO# * LOTENUEVO
           CIXXII% = XVALO(!CODIELEM)
           CODELEM$ = SAFETEXT$(!CODXELEM)
           NUOR% = XVALO(!NUORIT)
           
           'AGREGO A LA FORMULA PATRON EL NUEVO USO
10         RNDX$ = RANDSTRING$(8)
           If CantRegistros("FORMUPAT", "CODXCONJ = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
           '
           Call CreaRegistro("FORMUPAT", "CODXCONJ", RNDX$, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODICONJ", "CODXCONJ = '" & RNDX$ & "'", CI%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODIELEM", "CODXCONJ = '" & RNDX$ & "'", CIXXII%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODXELEM", "CODXCONJ = '" & RNDX$ & "'", CODEXT$(CIXXII%), RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "USOBRUTO", "CODXCONJ = '" & RNDX$ & "'", FORMATNUM$(USO#, "F16.5"), RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "NUORIT", "CODXCONJ = '" & RNDX$ & "'", NUOR%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODIEMPR", "CODXCONJ = '" & RNDX$ & "'", CodiEmp%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "MARBORRA", "CODXCONJ = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
           '
           Call ACTUREGISTRO("FORMUPAT", "CODXCONJ", "CODXCONJ = '" & RNDX$ & "'", CODEXT$(CI%), RAFE&, "NOMESS")
           '
        .MoveNext
        Loop
     End With
     FORMUTMP.Close
     Set FORMUTMP = Nothing
     '
     'BORRO LOS REGISTROS MARCADOS CON 1
     CONDI$ = "CodIConj = " & CI% & " AND MARBORRA = 1 "
     Call BORRAREGISTROS("FORMUPAT", CONDI$, RAFE&, "NOMESS")

''     'ACTUALIZO TABLA FORMUPAT
''     Dim FORMU As New ADODB.Recordset
''     FORMU.Open FILTSQL$("Select * from formupat where CodIConj = " & CI%), FLEXCONN, adOpenKeyset, adLockPessimistic
''     With FORMU
''        Do While Not .EOF
''             USO# = XVALO(!USOBRUTO)
''             !USOBRUTO = (USO# / LOTEANTERIOR) * LOTENUEVO
''             .Update
''             .MoveNext
''         Loop
''      End With
''      FORMU.Close
''      Set FORMU = Nothing
''      '
''      'Se borra el proceso de actualizar la formula unitaria por que en realidad
''      'la cantidad unitaria no cambia por que sigue manteniendo la relacion ni tampoco se
''      'esta cambiando nada de la estructura
      
      'BLANQUEAR LABEL DE FECHA DE REVISION
      CONDI$ = "CODINT = " & CI%
      Call ACTUREGISTRO("MASTER", "USUAPFOR", CONDI$, "PENDIENTE DE APROBACION", REG&)
      Call ACTUREGISTRO("Master", "FechApFor", CONDI$, "01/01/80", REG&)
      Call ACTUREGISTRO("Master", "Revi_Formu", CONDI$, "01/01/80", REG&)
      Label59 = ""
      Label59.ToolTipText = ""
      Label59.Refresh
      Command51.Visible = True
      '
      Call COMUNI("Al Modificar el Lote Económico se ha actualizado La Fomula Patrón.\Por Tal Motivo Debe Volver a Aprobarse.")
      '
End Sub

Sub Guardar_Rangos()
  Dim FORMU As ADODB.Recordset
  Set FORMU = New ADODB.Recordset
  CI% = CODINT%(Text1)
  Set FORMU = FLEXCONN.Execute(FILTSQL$("Select * from FormuPat where CodIConj = " & CI% & " ORDER BY NUORIT ASC"))
  If FORMU.BOF And FORMU.EOF Then
     Exit Sub
  End If
  '
   FORMU.MoveFirst
   REG& = 0
   Do While Not (FORMU.EOF)
     AAXX$ = REGBLAN("!RACEPTA")
     CODIELE% = XVALO(FORMU!CODIELEM)
     Unidad$ = UNIMED$(CODIELE%)
     If UCase(Unidad$) = "KG" Or UCase(Unidad$) = "GR" Or UCase(Unidad$) = "MG" Or UCase(Unidad$) = "MC" Then
       USO# = XVALO(FORMU!USOBRUTO)
       MINIMO# = XVALO(FORMU!Acepta_Min)
       maximo# = XVALO(FORMU!Acepta_Max)
       PREC# = XVALO(FORMU!Acepta_Prec)
       NUMITEM% = XVALO(FORMU!NUORIT)
       '
       Call REPLA(AAXX$, MKI(CODIELE%), 1, 2)
       Call REPLA(AAXX$, Unidad$, 3, 4)
       Call REPLA(AAXX$, MKD(USO#), 7, 8)
       Call REPLA(AAXX$, MKD(MINIMO#), 15, 8)
       Call REPLA(AAXX$, MKD(maximo#), 23, 8)
       Call REPLA(AAXX$, MKD(PREC#), 31, 8)
       Call REPLA(AAXX$, MKI(NUMITEM%), 39, 2)
       '
       REG& = REG& + 1
       Call GRAREG("!RACEPTA", AAXX$, REG&)
     End If
   FORMU.MoveNext
   Loop
   FORMU.Close
   Call SETULTREG("!RACEPTA", REG&)
   '
   CIXI% = CODINT(Text1)
   LOTE# = LOTESTAN(CIXI%)
   TITU$ = "Rango de Aceptación:  " & Text1 & "  x  " & LOTE# & "  " & UNIMED(CI%) & "  -  " & DESCRIT0(CI%)
   '
10 Respuesta% = VENTABU("RANGACEP/TITUPAN=" & TITU)
   If Respuesta% < 0 Then
      Exit Sub
   End If
   '
  CIXI% = CODINT(Text1)
  '
  FIN& = ULTREG("!RACEPTA")
  For REG& = 1 To FIN&
     AARR$ = REGBLAN("!RACEPTA")
     AARR$ = REGLEIDO("!RACEPTA", REG&)
     NUMITEM% = CVI(Mid(AARR$, 39, 2))
     CODIELE% = CVI(Mid(AARR$, 1, 2))
     Unidad$ = UNIMED(CODIELE%)
     USO# = CVD(Mid(AARR$, 7, 8))
     MINIMO# = CVD(Mid(AARR$, 15, 8))
     If MINIMO# > 0 Then
        COMUNI ("El Rango de Aceptación Minimo Debe ser Menor o Igual a Cero")
        GoTo 10
     End If
     '
     maximo# = CVD(Mid(AARR$, 23, 8))
     If maximo# < 0 Then
        COMUNI ("El Rango de Aceptación Maximo Debe ser Mayor o Igual a Cero")
        GoTo 10
     End If
     '
     PREC# = CVD(Mid(AARR$, 31, 8))
     If PREC# < 0 Then
        COMUNI ("La Precision debe ser  Mayor a Cero")
        GoTo 10
     End If
     CONDI$ = "CODICONJ = " & CIXI% & " And NUORIT = " & NUMITEM%
     Call ACTUREGISTRO("FORMUPAT", "ACEPTA_MIN", CONDI$, MINIMO#, REGI&)
     Call ACTUREGISTRO("FORMUPAT", "ACEPTA_MAX", CONDI$, maximo#, REGI&)
     Call ACTUREGISTRO("FORMUPAT", "ACEPTA_PREC", CONDI$, PREC#, REGI&)
  Next REG&
  
End Sub

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
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODINT = 1 + RecordCount
          .Update
          GoTo 45
        End If
        '
        On Error GoTo 0
        'On Error Resume Next
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !CODIGO = CODIG
        !CODINT = CIXI%
        !Status = STAT%
        !Descripcion = Descrip
        !UMEDIDA = UNIME
        !TIMATE = TIMA%
        !CRITREP = CRIT%
        !Monecos = MONEI%
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
    x$ = REGLEIDO$("ESTRUNUE", REG&)
    CICONJ% = CVI(Left$(x$, 2))
    DPS$ = Mid$(x$, 3, 120)
    REVIFORMU = CVI(Mid$(x$, 123, 2))
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
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
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
      x$ = REGLEIDO$("SECOPER", REG&)
      '
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
         CI% = CVI(Left$(x$, 2))
         RESECOX$ = x$
      !CODICONJ = CI%
      !CODXCONJ = CODEXT$(CI%)
      !NumeOper = CVI(Mid$(RESECOX$, 3, 2))
      !DescOper = TRIM$(Mid$(RESECOX$, 5, 48))
      !StatOper = 1
      !CodMaq0 = TRIM$(Mid$(RESECOX$, 53, 6))
      !CodMaq1 = TRIM$(Mid$(RESECOX$, 103, 6))
      !CodMaq2 = TRIM$(Mid$(RESECOX$, 109, 6))
      !CodMaq3 = TRIM$(Mid$(RESECOX$, 115, 6))
      !CodMaq4 = TRIM$(Mid$(RESECOX$, 182, 6))
      !CodMaq5 = TRIM$(Mid$(RESECOX$, 188, 6))
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
   Call Command21_CLICK
End Sub



Private Sub AtGral_Click()
Call OPLIMAS07.Command9_Click
End Sub

Private Sub ayu_Click()

End Sub

'Private Sub ayu_Click()
'Call Command4_Click
'End Sub

Private Sub BaLog_Click()
    Call Command24_Click
End Sub

Private Sub CapInv_Click()
    Call OPLIMAS07.Command15_Click
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

Private Sub Check3_Click()
  '
  MODIFIC%(0) = 1
  '
End Sub

Private Sub Check4_Click()
   MODIFIC%(0) = 1
End Sub


Private Sub Command47_Click(Index As Integer)
   'AGREGA / QUITA DOCUMENTO ANEXO O POE
   Command47(Index).Enabled = False
    
    Select Case Index
    'BOTON ADD
    Case 0: CARGARUTA
    'BOTON QUITAR
    Case 1
       If List2.ListIndex >= 0 Then
          List2.RemoveItem (List2.ListIndex)
       End If
   End Select
   '
   MODIFIC%(1) = 1
   '
99 Command47(Index).Enabled = True

End Sub
Sub CARGARUTA()
      'PERMITE AGREGAR NUEVOS POE'S
      '
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      RUTA$ = CD1.FileName
      If TRIM(RUTA$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", DESCRI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, RUTA, 32, 512)
      List2.AddItem TX$
      MODIFIC%(1) = 1
End Sub



Private Sub Command50_Click(Index As Integer)
   'CAMBIE A MATRIZ POR QUE NO PERMITE MAS OBJETOS
   Select Case Index
   Case 0
     COMIE = Text43.SelStart
     LOSEL = Text43.SelLength
     Text43.TEXT = Left$(Text43.TEXT, COMIE) + Mid$(Text43.TEXT, COMIE + LOSEL + 1)
   Case 1
     Clipboard.SetText Mid$(Text43.TEXT, Text43.SelStart + 1, Text43.SelLength)
   Case 2
     COMIE = Text43.SelStart
     LOSEL = Text43.SelLength
     Text43.TEXT = Left$(Text43.TEXT, COMIE) + Clipboard.GetText + Mid$(Text43.TEXT, COMIE + LOSEL + 1)
   Case 3
     Text43.SelStart = 0
     Text43.SelLength = Len(Text43.TEXT)
     Text43.SetFocus
   End Select
End Sub

Private Sub Command51_Click()
   Call ABREMASTER
   HOII% = HOY(HOS$)
   CONDI$ = "Codint=" & CODINT(Text1)
   Call ACTUREGISTRO("Master", "UsuApFor", CONDI$, USERNAME, REG&)
   Call ACTUREGISTRO("Master", "FechApFor", CONDI$, CVDAT(HOII%), REG&)
   Call ACTUREGISTRO("Master", "Revi_Formu", CONDI$, CVDAT(HOII%), REG&)
   Label59.Caption = FECHATEX$(FEREFOR%(CI%))
   Label59.ToolTipText = "Aprobada por " & SAFETEXT$(VALOBADA("Master", "UsuApFor", CONDI$, "NOMESS"))
   Label59.Refresh
   COMUNI ("Formula Aprobada")
   Command51.Visible = False
End Sub

Private Sub Command52_Click()
   '
   Command52.Enabled = False
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GABAL") Then GoTo 20
   '
   If CONTROL$("FACTURA", "SALIDA") <> "FELECT" Then ' TIENE QUE ESTAR CONFIGURADO FACT.ELEC.
      Call MENSERR(24, "Opción Solo Para Factura Electrónica")
      GoTo 99
   End If
   '
20 OPX% = ALTERNA%("  Actualización  \  Listados  ")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   '
   If OPX% = 1 Then  ' ACTUALIZACION
      CI% = CODINT(Text1) ' VALIDA QUE SE HAYA SELECCIONADO CODIGO
      If CI% < 1 Then
        Call MENSERR(24, "Código No Válido")
        GoTo 99
      End If
      '
      'TOMO LOS VALORES DE LA BASE DE DATOS PARA PRESENTAR EN EL OBJETO
      CONDI$ = "Codint = " & CI%
      TIPOX% = XVALO(VALOBADA("MASTER", "Bien_Cap", CONDI$, "NOMESS"))
      NOMMER$ = VALOBADA("MASTER", "Codigo_Ncm", CONDI$, "NOMESS")
      CODSEC$ = VALOBADA("MASTER", "Codigo_Sec", CONDI$, "NOMESS")
      COUNMED$ = VALOBADA("MASTER", "Pro_Umed", CONDI$, "NOMESS")
      COTIVA$ = VALOBADA("MASTER", "Iva_Id", CONDI$, "NOMESS")
      '
      PIXCAJA = XVALO(VALOBADA("MASTER", "PiezasXCaja", CONDI$, "NOMESS"))
      MT3CJA# = XVALO(VALOBADA("MASTER", "M3Caja", CONDI$, "NOMESS"))
      KNTCAJA# = XVALO(VALOBADA("MASTER", "KiloNetoCaja", CONDI$, "NOMESS"))
      KBTCAJA# = XVALO(VALOBADA("MASTER", "KiloBrutoCaja", CONDI$, "NOMESS"))
      '
      TipoEti% = XVALO(VALOBADA("MASTER", "Tamano_Etiqueta", CONDI$, "NOMESS"))
      '
     'CREO EL STRING PARA EL OBJETO
      Call REPLA(VDEF$, MKI$(TIPOX%), 1, 2) ' INFORMACION PARA EL OBJETO
      Call REPLA(VDEF$, NOMMER$, 3, 16)
      Call REPLA(VDEF$, CODSEC$, 19, 16)
      Call REPLA(VDEF$, COUNMED$, 35, 8)
      Call REPLA(VDEF$, COTIVA$, 43, 8)
      '
      Call REPLA(VDEF$, MKS$(PIXCAJA), 51, 4)
      Call REPLA(VDEF$, MKD$(MT3CJA#), 55, 8)
      Call REPLA(VDEF$, MKD$(KNTCAJA#), 63, 8)
      Call REPLA(VDEF$, MKD$(KBTCAJA#), 71, 8)
      Call REPLA(VDEF$, MKI$(TipoEti%), 79, 1)
      '
30    OBX$ = OBJPLA$("TIPROFAC", VDEF$) ' ABRE EL OBJETO
      '
      If OBX$ = "" Then GoTo 20
      '
      TIPOX% = CVI(Mid$(OBX$, 1, 2)) 'TOMO VALOR DE TIPO DEL OBJETO Y VALIDO
      If TIPOX% < 0 Or TIPOX% > 1 Then
         Call MENSERR(24, "Valor Ingresado en Tipo de Producto No Válido (0 o 1) ")
         VDEF$ = OBX$
         GoTo 30
      End If
      '
      'TOMO LOS DATOS DEL OBJETO
      NOMMER$ = Mid$(OBX$, 3, 16)
      CODSEC$ = Mid$(OBX$, 19, 16)
      COUNMED$ = Mid$(OBX$, 35, 8)
      COTIVA$ = Mid$(OBX$, 43, 8)
      '
      PIXCAJA = CVS(Mid$(OBX$, 51, 4))
      MT3CJA# = CVD(Mid$(OBX$, 55, 8))
      KNTCAJA# = CVD(Mid$(OBX$, 63, 8))
      KBTCAJA# = CVD(Mid$(OBX$, 71, 8))
      '
      TipoEti% = Asc(Mid$(OBX$, 79, 1))
      '
      'VALIDO LOS DATOS QUE NO SUPEREN EL LIMITE
      If ROUND(PIXCAJA) >= 10000000 Then Call COMUNI("Ingreso de Piezas X Caja Supera El Limite Permitido"):: GoTo 30
      If ROUND(MT3CJA#) >= 10000000 Then Call COMUNI("Ingreso de M3 X Caja Supera El Limite Permitido"): VDEF$ = OBX$: GoTo 30
      If ROUND(KNTCAJA#) >= 1000000000 Then Call COMUNI("Ingreso de Kg. X Caja Supera El Limite Permitido"): VDEF$ = OBX$: GoTo 30
      If ROUND(KBTCAJA#) >= 1000000000 Then Call COMUNI("Ingreso de Kg. Totales  Supera El Limite Permitido"): VDEF$ = OBX$: GoTo 30
      '
      'VALIDO EL TAMAÑO DE ETIQUETA
      If TipoEti% <> 0 Then
         If TipoEti% < 1 Or TipoEti% > 2 Then
            Call COMUNI("Valor Ingresado Para Tamaño de Etiqueta No Válido\Valores Admisibles '1' o '2'")
            VDEF$ = OBX$
            GoTo 30
         End If
      End If
      '
      'GRABO LOS DATOS INGRESADOS EN EL OBJETO
      Call ACTUREGISTRO("MASTER", "Bien_Cap", CONDI$, TIPOX%, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "Codigo_Ncm", CONDI$, NOMMER$, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "Codigo_Sec", CONDI$, CODSEC$, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "Pro_Umed", CONDI$, COUNMED$, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "Iva_Id", CONDI$, COTIVA$, REGAF&, "NOMESS")
      '
      Call ACTUREGISTRO("MASTER", "PiezasXCaja", CONDI$, PIXCAJA, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "M3Caja", CONDI$, MT3CJA#, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "KiloNetoCaja", CONDI$, KNTCAJA#, REGAF&, "NOMESS")
      Call ACTUREGISTRO("MASTER", "KiloBrutoCaja", CONDI$, KBTCAJA#, REGAF&, "NOMESS")
      '
      Call ACTUREGISTRO("MASTER", "Tamano_Etiqueta", CONDI$, TipoEti%, REGAF&, "NOMESS")
      '
   ElseIf OPX% = 2 Then ' LISTADO
      Call FINAL("?LIPROTIP") ' LISTADO DE TIPO PRODUCTO = 1 (BIEN CAPITAL)
   End If
   
99 Command52.Enabled = True
End Sub


Private Sub Command54_Click(Index As Integer)
    'CAMBIE A MATRIZ POR QUE NO PERMITE MAS OBJETOS

    Select Case Index
    Case 0
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      RUTA$ = CD1.FileName
      Label63.Caption = RUTA$
    Case 1
      Command54(Index).Enabled = False 'MODIFICADO OD-30/11/11 NO EXISTE COMMAND71
      Label63 = ""
      Command54(Index).Enabled = True
    End Select

End Sub

Private Sub Command55_Click()
   Call SELECHO("TAFAMIL")
   VALONUE$ = VALACT1$("TAFAMIL")
   If VALONUE$ <> "" Then
      Text48.TEXT = VALONUE$
   End If

End Sub

Private Sub Command56_Click()
   ' Muestra Objeto en el cual se carga el LTC por proveedor
   x$ = OBJPLA$("UPDATELTC", VDEF$)
   If x$ <> "" Then
     PRO% = CVI(Mid$(x$, 1, 2))
     LTC$ = TRIM$(Mid$(x$, 3, 5))
     CONDI$ = "PROVHABI = " & PRO% & " and (LeTiCom = '0' or LeTiCom is null)"
     Call ACTUREGISTRO("MASTER", "LeTiCom", CONDI$, LTC$, REG&)
     Call COMUNI("Actualizacion Finalizada. Cantidad de Registros Afectados = " & REG&)
   End If
End Sub

Private Sub Command57_Click()
    NMS0& = VALOBADA("MASTER", "MAX(CODINT)", "CODINT > 0")
    NMS& = NMS0&
    Call ABRECONEXION
    AGREGA& = 50000
    MsgBox "Hasta " & (NMS0& + AGREGA&)
    Do While NMS& < NMS0& + AGREGA&
       NMS& = NMS& + 1
       Call TRACE(20, NMS& - NMS0&, AGREGA&)
       Call CreaRegistro("MASTER", "CODINT", NMS&, "NOMESS/NOAUDIT")
       CONDI$ = "CODINT = " & NMS&
       CODGEN$ = TRIM$(Str$(NMS&)): While Len(CODGEN$) < 6: CODGEN$ = "0" + CODGEN$: Wend
       CODGEN$ = "K-" + CODGEN$
       Call ACTUREGISTRO("MASTER", "STATUS", CONDI$, 1, RAFE&, "NOAUDIT")
       Call ACTUREGISTRO("MASTER", "CODIGO", CONDI$, CODGEN$, RAFE&, "NOAUDIT")
       Call ACTUREGISTRO("MASTER", "DESCRIPCION", CONDI$, "Articulo Codigo " + CODGEN$, RAFE&, "NOAUDIT")
    Loop
    UPDATMASTER = 1
    Call GENECOMAS
End Sub

Private Sub Command58_Click()
    '
    Exit Sub
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "PHARMAFLEX") < 1 Then Exit Sub
    '
    Call ABRECONEXION
    Dim DTABLA$(256), DCAMPO$(256)
    '
    Screen.MousePointer = 11
    QTABLAS% = 0
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    QTABLAS% = 0
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       If InStr(UCase$(AUX!Name), "QUEUE") > 0 Then GoTo 29
       QTABLAS% = QTABLAS% + 1
       DTABLA$(QTABLAS%) = UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
    '
    For ITABLA% = 1 To QTABLAS%
       TAPRIN$ = DTABLA$(ITABLA%)
       Dim AUX1 As ADODB.Recordset
       Set AUX1 = New ADODB.Recordset
       strsql = "SELECT * FROM " & TAPRIN$
       AUX1.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       '
       With AUX1
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
              If UCase$(Left$(.Fields(KKI% - 1).Name, 8)) <> "VALOBADA" Then
                On Error Resume Next
                VATEX = ""
                VATEX = .Fields(KKI% - 1).Value
                On Error GoTo 0
                Call REPLATE(VATEX, DEVUX)
                If DEVUX <> VATEX Then
                  On Error Resume Next
                  .Fields(KKI% - 1).Value = DEVUX
                  .Update
                  On Error GoTo 0
                AAA = BBB
                End If
              End If
           Next KKI%
         .MoveNext
         Loop
       End With
       AUX1.Close
       Set AUX1 = Nothing
    Next ITABLA%
    '
    UPDATMASTER% = 1
    Call GENECOMAS
    '
    Screen.MousePointer = 1
    Call COMUNI("PREOCESO COMPLETO")
    '
End Sub
'
Sub REPLATE(TORI, TREP)
    TTTT = TORI
    TTTT = REPTE(TTTT, "SYNTH", "ESPECU")
    TTTT = REPTE(TTTT, "CUIDAFEM", "NOPREGNA")
    TTTT = REPTE(TTTT, "ABBOTT", "BEBOT")
    TTTT = REPTE(TTTT, "POVIDON", "BESILAN")
    TTTT = REPTE(TTTT, "CITRO", "NARANJ")
    TTTT = REPTE(TTTT, "ALELI", "BARBAMAM")
    TTTT = REPTE(TTTT, "KALA", "GLADIOL")
    TTTT = REPTE(TTTT, "OPADR", "EMAMM")
    TTTT = REPTE(TTTT, "BLIPACK", "TRICEPS")
    TTTT = REPTE(TTTT, "LEVONO", "PAPALE")
    TTTT = REPTE(TTTT, "PROGEST", "NOPREG")
    TTTT = REPTE(TTTT, "DIVA", "BELLADONNA")
    TTTT = REPTE(TTTT, "EQUIFEM", "TAMBAMACH")
    TTTT = REPTE(TTTT, "GESTODENO", "PREGNALOCK")
    TTTT = REPTE(TTTT, "TAMOXIFENO", "REXIPLACK")
    TTTT = REPTE(TTTT, "FLUTA", "GASCO")
    TTTT = REPTE(TTTT, "TRIDESTAN", "TETRASTOP")
    TTTT = REPTE(TTTT, "OPAGLOS", "NONAVID")
    TTTT = REPTE(TTTT, "LEVONOGEST", "FALLPREST")
    TTTT = REPTE(TTTT, "CELLAC", "MICROSEC")
    TTTT = REPTE(TTTT, "METROCEL", "YARDLOCK")
    TTTT = REPTE(TTTT, "ASTARTE", "BEPINTU")
    TTTT = REPTE(TTTT, "ISIS", "SPHI")
    TTTT = REPTE(TTTT, "ANAST", "NASOT")
    TTTT = REPTE(TTTT, "ANUBIS", "CHEOPS")
    TTTT = REPTE(TTTT, "CLIMA", "TEMPO")
    TTTT = REPTE(TTTT, "DROSPI", "TRASMA")
    TTTT = REPTE(TTTT, "KOLLIDON", "HEADMAS")
    TTTT = REPTE(TTTT, "TIBOLONA", "LESFEROA")
    TTTT = REPTE(TTTT, "MEGASTROL", "KILOTRAN")
    TTTT = REPTE(TTTT, "LIF", "MAGIC")
    TTTT = REPTE(TTTT, "GLANIQUE", "NEGROCUA")
    TTTT = REPTE(TTTT, "SECUR", "PERIC")
    TTTT = REPTE(TTTT, "SINTEMEL", "ARTIFIC")
    TTTT = REPTE(TTTT, "MENOTRIX", "MADUREX")
    TTTT = REPTE(TTTT, "PRICUL", "LASTHED")
    TTTT = REPTE(TTTT, "ASOFLUT", "BEXIL")
    TTTT = REPTE(TTTT, "ACTOSE", "BECALA")
    TTTT = REPTE(TTTT, "DASTONIL", "BEXINOT")
    TTTT = REPTE(TTTT, "SENAL", "JUVEN")
    TTTT = REPTE(TTTT, "AST", "IMB")
    TTTT = REPTE(TTTT, "PROTER", "NEGACA")
    TTTT = REPTE(TTTT, "PIRRO", "MESSO")
    TTTT = REPTE(TTTT, "PROMEL", "ANTSUG")
    TTTT = REPTE(TTTT, "KOLLI", "MEXA")
    TTTT = REPTE(TTTT, "CEPRA", "JIRRA")
    TTTT = REPTE(TTTT, "TRIACA", "UBALDO")
    TTTT = REPTE(TTTT, "CIPRO", "TREXI")
    TTTT = REPTE(TTTT, "GONDO", "PANFE")
    TTTT = REPTE(TTTT, "ASOFAR", "IKAFAR")
    TTTT = REPTE(TTTT, "BERNABO", "KARKISO")
    TTTT = REPTE(TTTT, "ELEA", "ALOU")
    TTTT = REPTE(TTTT, "INVESTI", "OUTNAK")
    TTTT = REPTE(TTTT, "LIBRA", "KILOX")
    TTTT = REPTE(TTTT, "VERDE", "AZULI")
    TTTT = REPTE(TTTT, "PELLIER", "CLANOR")
    TTTT = REPTE(TTTT, "SIDUS", "ARMEX")
    TREP = TTTT
'If InStr(TORI, "CUIDAFEM") > 0 Then
'AAA = BBB
'End If
    If TREP <> TORI Then
        AAA = BBB
    End If
End Sub
'
Function REPTE(TORI, BORI$, BREP$)
    '
    REPTE = TORI
    If Len(BORI$) > 0 Then
       If InStr(UCase$(BREP$), UCase$(BORI$)) < 1 Then
10        PPXX& = InStr(UCase$(SAFETEXT$(REPTE)), UCase$(BORI$))
          If PPXX& > 0 Then
             IZQUI$ = Left$(REPTE, PPXX& - 1)
             DERE$ = Mid$(REPTE, PPXX& + Len(BORI$))
             REPTE = IZQUI$ & BREP$ & DERE$
             GoTo 10
          End If
       End If
    End If
    '
End Function
'
Private Sub Command67_Click()
   Command67.Enabled = False
   Screen.MousePointer = 11
   
   If DERACCE%("SORTSTRU", "Ordenamiento de Estructura de Producto") < 2 Then
      GoTo 99
   End If

   For I% = 0 To List1.ListCount - 1 ' ORDENAMIENTO
     CI1$ = TRIM$(Mid$(List1.List(I%), 1, 16))
     For J% = I% + 1 To List1.ListCount - 1
       CI2$ = TRIM$(Mid$(List1.List(J%), 1, 16))
       If UCase$(CI1$) > UCase$(CI2$) Then
         COD1$ = List1.List(I%)
         COD2$ = List1.List(J%)
         List1.List(I%) = COD2$
         List1.List(J%) = COD1$
         CI1$ = CI2$
       End If
     Next J%
   Next I%
   MODIFIC%(2) = 1
99 Command67.Enabled = True
   Screen.MousePointer = 1

End Sub

Private Sub Command68_Click()

  If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub

  Call SELECHO("SELFECHA")
  VALONUE$ = VALACT1$("SELFECHA")
  If VALONUE$ <> "" Then
      Text83.TEXT = VALONUE$
  End If

End Sub

Private Sub Command69_Click()
  If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub

  Call SELECHO("SELFECHA")
  VALONUE$ = VALACT1$("SELFECHA")
  If VALONUE$ <> "" Then
      Text84.TEXT = VALONUE$
  End If

End Sub

Private Sub Command70_Click()
    '
    Command70.Enabled = False
    '
    'VALIDO QUE ESTE SELECCIONADO UNA OPERACION
    If XVALO(Label55) < 1 Then
       Call COMUNI("Debe Ingresar una Operación")
       GoTo 99
    End If
    '
10  OPX% = COMALTER%("Opciones de Trabajo:\\Herramientas \ Equipos Auxiliares")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    'PONGO EN 0 LOS VALORES NULL
    Call PoneEnCeroNull("HERRAOPER", "Tipo_Herr", REGAF&, "NOMESS")
    '
    'ASIGNO VALOR A VARIABLE SEGUN OPCION ELEGIDA
    If OPX% = 1 Then
       TIPOHERRAM% = 0 'HERRAMENTALES
       ARCHIX$ = "HERRAOP" ' SOLO PARA LA VENTANA POR QUE ESTRUCTURA ES LA MISMA
    Else
       TIPOHERRAM% = 1 ' EQUIPOS AUXILIARES
       ARCHIX$ = "EQUIOPER"
    End If
    '
    Call SETULTREG("!HERRAOP", 0)
    CI% = CODINT(Text1)
    COPE% = XVALO(Label55)
    '
    'VUELCO LA INFORMACION DE LA BASE A LA VENTANA
    J& = 0
    SQLX$ = "SELECT * FROM HERRAOPER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODINT =" & CI%
    SQLX$ = SQLX$ + " AND Cod_Oper = " & COPE%
    SQLX$ = SQLX$ + " AND Tipo_Herr = " & TIPOHERRAM%
    Dim HERRAOPTMP As ADODB.Recordset
    Set HERRAOPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With HERRAOPTMP
      Do While Not .EOF
        CODIG$ = SAFETEXT$(!Cod_Herramental)
        CANTI = XVALO(!Cantidad)
        If CODIG$ <> "" Then
        Call REPLA$(x$, CODIG$, 1, 16)
        Call REPLA$(x$, MKS(CANTI), 17, 4)
        J& = J& + 1
        Call GRAREG("!HERRAOP", x$, J&)
        End If
        .MoveNext
      Loop
    End With
    '
    'PRESENTA VENTANA DE CARGA
20  VTB% = VENTABU%(ARCHIX$)
    If VTB% < 0 Then GoTo 10
    '
    'MARCO LOS ITEMS PARA LUEGO BORRAR
    CONDI$ = "CODINT =" & CI% & " AND  Cod_Oper = " & COPE% & " AND Tipo_Herr = " & TIPOHERRAM%
    Call ACTUREGISTRO("HERRAOPER", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    'AGREGO LA INFORMACION DE LA VENTANA A LA TABLA
    SQLX$ = "SELECT * FROM HERRAOPER "
    Dim HERRAOPERTMP As ADODB.Recordset
    Set HERRAOPERTMP = New ADODB.Recordset
25      On Error Resume Next
    HERRAOPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0

    '
    'RECORRO LA VENTANA
    For U& = 1 To ULTREG&("!HERRAOP")
        x$ = REGLEIDO$("!HERRAOP", U&)
        CODIG$ = TRIM$(Left$(x$, 16))
        CANTI = CVS(Mid$(x$, 17, 4))
        If CODIG$ <> "" Then
           If CANTI < 1 Then Call COMUNI("Falta Ingresar Cantidad para Código: " & CODIG$): GoTo 20
            With HERRAOPERTMP
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                !CODINT = CI%
                !Cod_Herramental = CODIG$
                !Cantidad = CANTI
                !Cod_Oper = COPE%
                !Tipo_Herr = TIPOHERRAM%
                !MarBorra = 0
                .Update
            End With
        End If
    Next U&
    HERRAOPERTMP.Close
    Set HERRAOPERTMP = Nothing
    '
    'BORRO LOS REGISTROS ANTERIOERES QUE ESTABAN MARCADOS CON MARBORRA = 1
    CONDI$ = "MARBORRA = 1 "
    Call BORRAREGISTROS("HERRAOPER", CONDI$, REGAF&, "NOMESS")
    '
    Call CIERRARCH("!HERRAOP")
    '
    GoTo 10
99  Command70.Enabled = True

End Sub


'Private Sub Command71_Click()
'   Command71.Enabled = False
'   Label63 = ""
'   Command71.Enabled = True
'End Sub

Private Sub Label18_DblClick()
   Label71_Click
End Sub


Private Sub Label27_DblClick()
  
  If PORPATRON% < 1 Then
       Exit Sub
  End If
  '
  CIXI% = CODINT(Text1)
  If CIXI% <= 0 Then
      Exit Sub
  End If
  LOTE# = LOTESTAN(CIXI%)
  If LOTE# < 1 Then
     LOTE# = 1
   End If
  TITU$ = "Form. Patron:  " & Text1 & "  x  " & LOTE# & "  " & UNIMED(CI%) & "  -  " & DESCRIT0(CI%)
  Call SETULTREG("!MATEROF", 0)
  RESP% = COMALTER%("Opciones:\\ABM de Formula Patron \ Rangos de Aceptacion")
  If RESP% < 1 Then GoTo 99
  If RESP% = 2 Then
     Call Guardar_Rangos
     GoTo 99
  End If
  Dim FORMU As ADODB.Recordset
  Set FORMU = New ADODB.Recordset
  Set FORMU = FLEXCONN.Execute(FILTSQL$("Select * from FormuPat where CodIConj = " & CI% & " ORDER BY NUORIT ASC"))
  
  If FORMU.BOF And FORMU.EOF Then
     GoTo 10
  End If
   
  'Completa MATEROF para presentar en pantalla
     FORMU.MoveFirst
     REG& = 0
   Do While Not (FORMU.EOF)
     AAXX$ = REGBLAN("MATEROF")
     CODIELE% = XVALO(FORMU!CODIELEM)
     Unidad$ = UNIMED$(CODIELE%)
     USO# = XVALO(FORMU!USOBRUTO)
     Operacion% = XVALO(FORMU!Operacion)
     '
     Call REPLA(AAXX$, MKI(CODIELE%), 1, 2)
     Call REPLA(AAXX$, Unidad$, 3, 2)
     Call REPLA(AAXX$, MKD(USO#), 9, 8)
     Call REPLA(AAXX$, MKI(Operacion%), 29, 2)
     '
     REG& = REG& + 1
     Call GRAREG("!MATEROF", AAXX$, REG&)
     FORMU.MoveNext
   Loop
     '
10   Respuesta% = VENTABU("FORPATRO/TITUPAN=" & TITU)
     If Respuesta% < 0 Then
        Exit Sub
     End If
     '
     'Decide si se vuelve a Aprobar la Formula Patron
     If DERACCE%("MODFOPAT", "Modificar Formula Patrón") < 2 Then
        GoTo 99
     End If
     '
     TTXX$ = "Desea Registrar la Formula Patrón Modificada?"
     If CONTROL("FORMULAS", "REQAPRO") = "SI" Then
        TTXX$ = TTXX$ + "\(La Fómula Deberá ser Nuevamente Aprobada)"
     End If
     RESMOD% = COMALTER(TTXX$ + "\\No, Deja Anterior\Si, Registrar Modificada")
     If RESMOD% <> 2 Then
          GoTo 99
     End If
     '
     CONDI$ = "CODINT = " & CIXI%
     Call ACTUREGISTRO("MASTER", "USUAPFOR", CONDI$, "PENDIENTE DE APROBACION", REG&)
     Call ACTUREGISTRO("Master", "FechApFor", CONDI$, "01/01/80", REG&)
     Call ACTUREGISTRO("Master", "Revi_Formu", CONDI$, "01/01/80", REG&)
     Label59 = ""
     Label59.ToolTipText = ""
     Label59.Refresh
     Command51.Visible = True
     '
     ' PONE MARCA DE BORRADO A PRE-EXISTENTES
     CONDI$ = "CodIConj = " & CI%
     Call ACTUREGISTRO("FORMUPAT", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
       
     'Recorre archivo Materof para actualizar Tabla Formupat
     FORMU.Close
     FORMU.Open FILTSQL$("Select * from formupat where CodIConj = " & CI%), FLEXCONN, adOpenKeyset, adLockPessimistic
     REGACTI& = 1
     REGFINAL% = ULTREG("!Materof")
     Do While REGACTI& <= REGFINAL%
          AAXX$ = REGLEIDO("!MATEROF", REGACTI&)
          CODIELE% = CVI(Mid$(AAXX$, 1, 2))
          Unidad$ = Mid$(AAXX$, 3, 2)
          USO# = CVD(Mid$(AAXX$, 9, 8))
          OPERA% = CVI(Mid$(AAXX$, 29, 2))
          '
          FORMU.AddNew   ' YA CONVERTIDO
               On Error Resume Next
          FORMU!CodiEmpr = CodiEmp%
               On Error GoTo 0
           FORMU!CODICONJ = CIXI%
           FORMU!CODXCONJ = CODEXT$(CIXI%)
           FORMU!CODIELEM = CODIELE%
           FORMU!CODXELEM = CODEXT$(CODIELE%)
           FORMU!USOBRUTO = USO#
           FORMU!Operacion = OPERA%
           FORMU!NUORIT = REGACTI&
           FORMU!MarBorra = 0
          FORMU.Update
          
          REGACTI& = REGACTI& + 1
      Loop
      FORMU.Close
      Set FORMU = Nothing
      '
      CONDI$ = "CodIConj = " & CI% & " AND MARBORRA = 1 "
      Call BORRAREGISTROS("FORMUPAT", CONDI$, RAFE&, "NOMESS")
      '
      ' PONE MARCA DE BORRADO A PRE-EXISTENTES EN FORMULAS
      CONDI$ = "CodIConj = " & CI%
      Call ACTUREGISTRO("FORMULAS", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
      '
      Dim FormuX As ADODB.Recordset
      Set FormuX = New ADODB.Recordset
      FormuX.Open FILTSQL$("select * from Formulas where codiconj = " & CIXI%), FLEXCONN, adOpenKeyset, adLockPessimistic
      '
      Dim FORPAT As ADODB.Recordset
      SQLX$ = "SELECT CODIELEM, SUM(USOBRUTO) AS SUBRUTO from FORMUPAT WITH(NOLOCK) WHERE CODICONJ = " & CIXI% & " GROUP BY CODIELEM"
      Set FORPAT = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      If Not (FORPAT.BOF And FORPAT.EOF) Then
         NumOrd% = 0
          
          FORPAT.MoveFirst
         Do While Not FORPAT.EOF
            NumOrd% = NumOrd% + 1
            ' OJO !!!! - EL USO DE LA FORMULA UNITARIO SE ALMACENA POR MILLAR
            USO# = ((XVALO(FORPAT!SUBRUTO)) / LOTE#)
            CODIELE% = FORPAT!CODIELEM
            FormuX.AddNew   ' YA CONVERTIDO
                  On Error Resume Next
            FormuX!CodiEmpr = CodiEmp%
                  On Error GoTo 0
            FormuX!CODICONJ = CIXI%
            FormuX!CODXCONJ = CODEXT$(CIXI%)
            FormuX!CODIELEM = CODIELE%
            FormuX!CODXELEM = CODEXT$(CODIELE%)
            FormuX!USOBRUTO = USO#
            FormuX!NUORIT = NumOrd%
            FormuX!MarBorra = 0
            FormuX.Update
          FORPAT.MoveNext
          Loop
      End If
      FormuX.Close
      Set FormuX = Nothing
      '
      CONDI$ = "CodIConj = " & CIXI% & " AND MARBORRA = 1 "
      Call BORRAREGISTROS("FORMULAS", CONDI$, RAFE&, "NOMESS")
     
      Call Text1_Change
99
End Sub



Private Sub Label58_Change()
   '
   UVALUE% = XVALO(Label58)
   If UVALUE% = 0 Then Label61 = "Trazable"
   If UVALUE% = 1 Then Label61 = "Trazable Lote"
   If UVALUE% = 2 Then Label61 = "Trazable Serie"
   '
End Sub

'
Private Sub label58_Click()
  '
  MODIFIC%(0) = 1
  '
  UVALUE% = XVALO(Label58)
  UVALUE% = (UVALUE% + 1) Mod 3
  Label58 = FORMATNUM$(UVALUE%, "I1")
  '
End Sub


Private Sub cmdConvertir_Click()
   Call ABRECONEXION
End Sub

Private Sub CodArt_Click()
   Call OPLIMAS07.Command4_Click
End Sub

Private Sub Command17_Click()
   CICONJ% = CODINT%(Text1)
   If CantRegistros("FORMUPAT", "CodIConj = " & CICONJ%) > 0 Then
      Call MENSERR(24, "Debe Actualizar Unicamente\por Fórmula Patrón.")
      Exit Sub
   End If
   Frame3.Visible = False
   Call COMMAND31_CLICK
End Sub

Private Sub Command43_Click()
   Command43.Enabled = False
   Call Text22_DblClick
   Command43.Enabled = True
End Sub

Private Sub Command44_Click()
   Command44.Enabled = False
   Call Text23_DblClick
   Command44.Enabled = True
End Sub

'Private Sub Command47_Click()
'   Text43.SelStart = 0
'   Text43.SelLength = Len(Text43.TEXT)
'   Text43.SetFocus
'End Sub



Private Sub Command53_Click()
   '
   ' instertar operacion alternativa
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   If List3.ListCount >= MAXECO% Then
      Call MENSERR(24, "Imposible Agregar Alternativa")
      Exit Sub
   End If
   '
   ILISTAX& = List3.ListIndex
   If ILISTAX& < 0 Or ILISTAX& >= List3.ListCount Then
      Call MENSERR(24, "Imposible Agregar Alternativa")
      Exit Sub
   End If
   '
   NOPEP% = XVALO(Left$(List3.List(ILISTAX&), 4))
   DOPEP$ = DESCRIOP$(CI%, NOPEP%)
   NOPY% = NOPEP%
   '
   For UI& = List3.ListCount - 1 To ILISTAX& Step -1
     If CVI(Mid$(RESECO$(UI&), 3, 2)) <= NOPY% Then
       ALTERX% = 1 + Asc(Mid$(RESECO$(UI&), 88, 1))
       ILISTAX& = UI& + 1
       GoTo 30
     End If
     RESECO$(UI& + 1) = RESECO$(UI&)
     DESOPE$(UI& + 1) = DESOPE$(UI&)
   Next UI&
   Call MENSERR(24, "Imposible Agregar Alternativa")
   Call Command2_Click    ' restaurar valores
   Exit Sub
   '
30 RESECO$(ILISTAX&) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(40, 0) 'para acomodar hoja de procesos DE FABRICACION
   Mid$(RESECO$(ILISTAX&), 88, 1) = Chr$(ALTERX%)
   DESOPE$(ILISTAX&) = ""
   NOPEX$ = Left$(AJUSTD$(Str$(NOPY%), 4) + "-" + TRIM$(Str$(ALTERX%)) + "   ", 7)
   DOPE$ = AJUSTI$("(Alt) " + DOPEP$, 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "  " + DOPE$ + " " + REGOX$
   MODIFIC%(3) = 1
   List3.AddItem SECO$, ILISTAX&

   List3.ListIndex = ILISTAX&
   Text5.SetFocus
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If EXISTE%("TRAZAB07.EXE") Then
        OPX% = ALTERNA%("Costo de Reposición por Lote\Otros Accesos Directos")
        If OPX% = 1 Then
            ACOSLOTE.Show 1
          ElseIf OPX% = 2 Then
            Call ACCDIR("AMASTOK")
        End If
        GoTo 99
   End If
   '
   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "ALEJPA") Then
      Call CARGA_CICLOS
   Else
      Call ACCDIR("AMASTOK")
   
   End If

99 Command6.Enabled = True
End Sub

Sub CARGA_CICLOS()
    '
    Call CREACAMPO("", "MASTER", "CANT_PZA_CICLO", 4, 0)
    Call CREACAMPO("", "MASTER", "CANT_EQUIV_SEGUNDOS", 4, 0)
    Call CREACAMPO("", "MASTER", "TMPO_CICLO_REAL", 4, 0)
    '
    'VALIDO QUE HAYA UN CODIGO ACTIVO
    CI% = XVALO(Label28)
    If CI% < 1 Then
      Call COMUNI("Código Inexistente o No Válido")
      Exit Sub
    End If
    '
    CONDI$ = "CODINT = " & CI%
    '
    CANT_PZA_CICLOX = XVALO(VALOBADA("MASTER", "CANT_PZA_CICLO", CONDI$, "NOMESS"))
    CANT_EQUIV_SEGUNDOSX = XVALO(VALOBADA("MASTER", "CANT_EQUIV_SEGUNDOS", CONDI$, "NOMESS"))
    TMPO_CICLO_REALX = XVALO(VALOBADA("MASTER", "TMPO_CICLO_REAL", CONDI$, "NOMESS"))
    '
    'PRESENTO OBJETO (VALOR DE TMPO CICLO REAL 36 POR DEFECTO)
    VDEF$ = ""
    Call REPLA(VDEF$, MKS$(CANT_PZA_CICLOX), 1, 4)
    Call REPLA(VDEF$, MKS$(CANT_EQUIV_SEGUNDOSX), 5, 4)
    Call REPLA(VDEF$, MKS$(TMPO_CICLO_REALX), 9, 4)
    If TMPO_CICLO_REALX < 0.5 Then Call REPLA(VDEF$, MKS$(36), 9, 4)
    '
    OBX$ = OBJPLA$("CICLOTIEM", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    'VDEF$ = OBX$
    '
    CANT_PZA_CICLOX = CVS(Mid$(OBX$, 1, 4))
    CANT_EQUIV_SEGUNDOSX = CVS(Mid$(OBX$, 5, 4))
    TMPO_CICLO_REALX = CVS(Mid$(OBX$, 9, 4))
    '
    Call ACTUREGISTRO("MASTER", "CANT_PZA_CICLO", CONDI$, FORMATNUM$(CANT_PZA_CICLOX, "F10.0"), REG&, "NOMESS")
    Call ACTUREGISTRO("MASTER", "CANT_EQUIV_SEGUNDOS", CONDI$, FORMATNUM$(CANT_EQUIV_SEGUNDOSX, "F10.0"), REG&, "NOMESS")
    Call ACTUREGISTRO("MASTER", "TMPO_CICLO_REAL", CONDI$, FORMATNUM$(TMPO_CICLO_REALX, "F10.0"), REG&, "NOMESS")
    '
End Sub

Private Sub Command61_Click()
   '
   Command61.Enabled = False
   '
   If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
      GoTo 99
   End If
   '
   CODD$ = TRIM$(Text1.TEXT)
   CIXX% = CODINT%(CODD$)
   If CIXX% <= 0 Then
      GoTo 99
   End If
   '
   J& = 0
   For U& = 1 To ULTREG&("LISTAS")
       NLIST$ = TRIM$(Mid$(REGLEIDO$("LISTAS", U&), 2, 32))
       If NLIST$ = "" Then NLIST$ = "Lista Numero " + TRIM$(Str$(U&))
       y$ = REGBLAN$("ACPREVEN")
       CLIS$ = TRIM$(Str$(U&))
       While Len(CLIS$) < 3: CLIS$ = "0" + CLIS$: Wend
       APREC$ = "LIPRE" + CLIS$
       PRLST# = PRELISTA#(APREC$, CIXX%)
       If PRLST# >= 0.00005 Then
          Call REPLA(y$, MKD$(PRLST#), 33, 8)
       End If
       Call REPLA(y$, NLIST$, 1, 32)
       J& = J& + 1
       Call GRAREG("!ACPREVEN", y$, J&)
   Next U&
   Call SETULTREG("!ACPREVEN", J&)
   '
   VTB% = VENTABU%("ACPREVEN")
   If VTB% >= 0 Then
     If DERACCE%("ACPREVEN", "Actualizar Precios de Venta") < 1 Then
       GoTo 99
     End If
     '
     Dim PRECILIS As ADODB.Recordset
     '
     Call APERBASDAT("PEDCLI")
     For NUMLIS& = 1 To ULTREG&("!ACPREVEN")
        y$ = REGLEIDO$("!ACPREVEN", NUMLIS&)
        PRLST# = CVD(Mid$(y$, 33, 8))
        ABUSQUE$ = "WHERE COD_LISTA = " & NUMLIS& & " AND COD_PIEZA = " & CIXX%
        '
        ' SI EL PRECIO ES CERO, LO ELIMINA DE LA LISTA
        If PRLST# < 0.00005 Then
           On Error Resume Next
            FLEXCONN.Execute (FILTSQL$("DELETE PREVENTA " + ABUSQUE$))
           On Error GoTo 0
           GoTo 39
        End If
        '
        'CONSULTA SI EXISTE EL ARTICULO EN LA LISTA DE PRECIOS
        'SI EXISTE ACTUALIZA EL PRECIO
        SQLX$ = "SELECT * FROM PREVENTA " + ABUSQUE$
        Set PRECILIS = New ADODB.Recordset
        PRECILIS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With PRECILIS
           If Not .BOF And Not .EOF Then
             .Close
             SQLX$ = "UPDATE PREVENTA SET Precio_Unitario = " & FORMATNUM$(PRLST#, "F15.6") & " " & ABUSQUE$
             FLEXCONN.Execute SQLX$
             GoTo 39
           End If
        End With
        PRECILIS.Close
        '
        SQLX$ = "SELECT MAX(NUM_ORDEN) AS MaxNumOrden FROM PREVENTA "
        SQLX$ = SQLX$ + "WHERE COD_LISTA = " & NUMLIS& & " "
        Set PRETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        MAORD& = 0
        On Error Resume Next
        MAORD& = PRETEMP!MaxNumOrden
        On Error GoTo 0
        '
        SQLX$ = "SELECT * FROM PREVENTA " + ABUSQUE$
        Set PRECILIS = New ADODB.Recordset
        PRECILIS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With PRECILIS
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           !Cod_Lista = NUMLIS&
           !Cod_Pieza = CIXX%
           !Num_Orden = MAORD& + 1
           !Desc_Pieza = Mid$(y$, 3, 42)
           !Denom_Pieza = DESCRIT0$(CIXX%)
           !Carac_Control = Asc(Mid$(y$, 45, 1))
           !Num_Original = Asc(Mid$(y$, 46, 16))
           !Precio_Unitario = PRLST#             'CVD(Mid$(Y$, 62, 8))
           !Cant_Limite = CVD(Mid$(y$, 70, 2))
           !Porc_Limite = CVD(Mid$(y$, 72, 1))
           .Update
        End With
        '
        PRECILIS.Close
        Set PRECILIS = Nothing
        '
39   Next NUMLIS&
   End If
   '
   Call CIERRARCH("*.*")
   '
99 Command61.Enabled = True
   '
End Sub

   '
Private Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("PLANEL07/" + TRIM$(Text1), "Planera Electrónica")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   CIXI% = CODINT%(Text1)
   If CIXI% > 0 Then
     Call FILTERFILE("MARCAMOD", "", 1, MKI$(CIXI%))
     VTB% = VENTABU%("MARCAMOD")
     If VTB% >= 0 Then
       If DERACCE%("ACTAMARC", "Actualizar Tabla de Marcas y Numeros de Parte") = 2 Then
         Call FILTERUPDATE("MARCAMOD", "", 1, MKI$(CIXI%))
       End If
     End If
   End If
   Command8.Enabled = True
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
Call OPLIMAS07.Command3_Click
End Sub

Private Sub CosRan_Click()
  Call Command50_Click(0)
End Sub

Private Sub DesCam_Click()
Call Command2_Click
End Sub

Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Label10(2) = CONTROL$("AMASTO", "ATRIPAR1")
      If Label10(2) = "" Then Label10(2) = ". Atributo 'A':"
      If Left$(Label10(2), 1) = "." Then Label10(2) = " " + Mid$(Label10(2), 2)
    Label10(3) = CONTROL$("AMASTO", "ATRIPAR2")
      If Label10(3) = "" Then Label10(3) = ". Atributo 'B':"
      If Left$(Label10(3), 1) = "." Then Label10(3) = " " + Mid$(Label10(3), 2)
      '
    '\\\OT-05-0633-WAR-22/09/05///
    PORLARGO% = 0
    Label10(20).Visible = False
    Label10(21).Visible = False
    Label10(22).Visible = False
    Label10(23).Visible = False
    Label36.Visible = False
    Label42.Visible = False
    Text41.Visible = False
    Text42.Visible = False

    RTA1$ = CONTROL$("AMASTO", "FORMULA")
    If UCase$(TRIM$(RTA1$)) = "CANTIDAD" Then
      '
      Label10(20).Visible = True ' cantidad
      Label10(21).Visible = True
      Label10(22).Visible = True
      Label10(23).Visible = True
      Label36.Visible = True
      Label42.Visible = True
      Text41.Visible = True
      Text42.Visible = True
      '
      Label4.Height = 250
      Label4.Top = 1130
      Label10(17).Top = 1130
      Frame13.Visible = True
      '
      Label40.Height = 250
      Label40.Top = 1130
      Label10(12).Top = 1130
      Frame16.Visible = True
      '
      PORLARGO% = 1
      Text32.BackColor = Label20.BackColor
      '
    End If
    '\\\OT-05-0633-WAR-FIN///
  
    Call APLICACIONSKINS(Me, Picture1)
    Call CENTRAFORM(Me)
    '
    'SI VIENE DE ANACOS09 (JYM)
    'SE REEMPLAZA "|" PR "/" POR QUE CUANDO LO PASA DE ANACOS09 LO REALILZA A LA INVERSA
    'POR QUE SI VIENE ORIGINALMENTE CON EL "/" LO TRUNCA HASTA AHI.
    If Left$(APLINVO$, 8) = "ANACOS09" Then
        CODX$ = TRIM$(Mid$(APLINVO$, 9))
        CODX$ = REPLACAR$(CODX$, "|", "/")
        If CODINT%(CODX$) > 0 Then
            Text1 = CODX$
       End If
    End If
    '
    If App.EXEName = "AMASTO07" Then
       Show
       Refresh
    End If
    ' Call VERJOBID
    '
    MAXECO% = 1024
    Screen.MousePointer = 11
    '
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    If HOY(HOS$) <= FECHANUM("01/08/10") Then Command57.Visible = True
    '
    Call RecuStoMinLotEco 'Recuperacion de Stock Minimo y Lote Economico menores a 0
    '
    Command7.Enabled = False
    If EXISTE%("PLANEL07.EXE") Then Command7.Enabled = True
    '
    Call CARTAVALISTO
    Call ACVEPROINT
    '
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    
    Screen.MousePointer = 1
    '
    If UCase$(APLINVO$) = "NEW" Then
      Call Command21_CLICK
      Call Command3_Click
    End If
    '
    CONDI$ = "CODINT=0"
    If CantRegistros("MASTER", CONDI$) > 0 Then Call BORRAREGISTROS("MASTER", CONDI$, REG&)
    CONDI$ = "COD_INTE=0"
    If CantRegistros("LOTINGRE", CONDI$) > 0 Then Call BORRAREGISTROS("LOTINGRE", CONDI$, REG&)
    '
    If CONTROL$("", "IDSECUN") = "ALIAS" Then
      Label10(41) = "Alias"
''''      Text45.FontSize = 8 COMENTADA ESTAS 3 LINEAS PARA QUE NO APAREZCA EL BOTON M
''''      Text45.Width = 2440
''''      Text45.FontSize = 11
'      Text3.FontSize = 5
'      Text3.Height = 270
'      Text3.FontSize = 9
      'Frame1.Visible = True
''''      Command8.Visible = True  COMENTADA ESTAS 3 LINEAS PARA QUE NO APAREZCA EL BOTON M
   End If
    '
    '
    HABIPORMI% = 0
    If CONTROL$("AMASTO", "PORMILES") = "SI" Then
      HABIPORMI% = 1
    End If
    '\\\OT-06-0252-WAR-26/06/06///
    Label47(2).ToolTipText = ""
    If EXISTE%(LUDAT$ + "TABLATMU.RFS") > 0 Then
      If ULTREG&("TABLATMU") > 0 Then
        EXISTMU% = 1
        Label47(2).ToolTipText = "Doble-Click para Abrir Tiempos Variables por Lote"
      End If
    End If
    '\\\OT-06-0252-WAR-FIN///
    '
    LOMA% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
    If LOMA% > 15 Then Text1.MaxLength = LOMA%
    '
    'Verifica si utiliza descripcion larga o descripcion de codigo y alias
    If CONTROL("MASTER", "DESLARGA") <> "SI" Then
        Text38.Visible = True
'       Frame1.Visible = True
'       Frame1.Height = 435
'       Frame1.Left = 1155
'       Frame1.Top = 880
'       Frame1.Width = 5580
    End If
    '
    ' Verifica si presenta o no formula Patron
    PORPATRON% = 0
    If CONTROL("MASTER", "FORPATRO") = "SI" Then
      PORPATRON% = 1
      Label27.ToolTipText = "Doble Clic Para Ver Formula Patron"
      Visible = True
      Screen.MousePointer = 11
      Call ACTUFECHAPRFOR
      Screen.MousePointer = 1
    End If
    '
    If CONTROL("MASTER", "M2PORUN") = "SI" Then
        Label10(4) = "M2/ Unid."
    End If
    '
    Call APERBASDAT("MASTER")
    '
End Sub
'
Sub ACTUFECHAPRFOR()
    '
    Call ABRECONEXION
    Call ABREMASTER
    SQLX$ = "select distinct MASTER.codint AS CDXX, MASTER.fechapfor from master "
    SQLX$ = SQLX$ + " inner join formupat on master.codint = formupat.codiconj "
    SQLX$ = SQLX$ + " where master.fechapfor is null"
    Dim LISTRANTMP As ADODB.Recordset
    Set LISTRANTMP = New ADODB.Recordset
    Set LISTRANTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    FFEE$ = FETEXCOR1$(FECHANUM("01/01/80"))
    With LISTRANTMP
      Do While Not .EOF
        CCII% = XVALO(!CDXX)
        CONDI$ = "CODINT = " & CCII%
        Call ACTUREGISTRO("MASTER", "FECHAPFOR", CONDI$, FFEE$, REG&, "NOMESS")
        .MoveNext
      Loop
    End With
    LISTRANTMP.Close
    Set LISTRANTMP = Nothing
    '
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   IMPLOSI07.Hide
   On Error GoTo 0
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
   '
   Command1.Enabled = False
   If MODOEDIT% = 1 Then Call MENSAEDIT: GoTo 99
   GRABATODO% = 1
   Call GRADATITEM
   Call ACTECOMAS
   GRABATODO% = 0
   '
99 Command1.Enabled = True
End Sub

Private Sub Command10_Click()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Command11_Click()
   '
   If EXISTE(LUDAT$ + "\TABTIMB.RFS") > 0 Then
      ARCHIVO$ = "TABTIMB"
     Else
      ARCHIVO = "TATIMAT"
   End If
   '
   Call SELECHO(ARCHIVO$)
   VALONUE$ = VALACT1$(ARCHIVO$)
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
   '
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
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(XVALO(Text10.TEXT))))
   End If
End Sub

Private Sub Command15_Click()
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
    CICONJ% = CODINT%(Text1)
    If CantRegistros("FORMUPAT", "CodIConj = " & CICONJ%) > 0 Then
       Call MENSERR(24, "Debe Actualizar Unicamente\por Fórmula Patrón.")
       Exit Sub
    End If
    '
    Frame3.Visible = False
    Text31.TEXT = Left$(List1.TEXT, 16)
    CDYY% = CODINT%(Text31.TEXT)
    Label40.Caption = DESCRIT0$(CDYY%)
    Label41.Caption = UNIMED$(CDYY%)
    
    
    Text32.TEXT = TRIM$(Mid$(List1.TEXT, 49, 12))
    Text27.TEXT = TRIM$(Mid$(List1.TEXT, 61, 4))
    Text39 = TRIM$(Mid$(List1.TEXT, 129))
    '\\\OT-05-0633-WAR-23/09/05///
    Text41.TEXT = TRIM$(Mid$(List1.TEXT, 89, 9))
    Label42 = FORMATNUM$(PESMETRO(CDYY%), "F7.3")
    Text42.TEXT = TRIM$(Mid$(List1.TEXT, 111, 6))
    Label36 = TRIM$(Mid$(List1.TEXT, 119, 8))
    '\\\OT-05-0633-WAR-FIN///
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command18.Enabled = False
    Command19.Enabled = False
    '
    MODOEDIT% = 1
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
    CICONJ% = CODINT%(Text1)
    If CantRegistros("FORMUPAT", "CodIConj = " & CICONJ%) > 0 Then
       Call MENSERR(24, "Debe Actualizar Unicamente\por Fórmula Patrón.")
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
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Command21_CLICK()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   If DERACCE%("ALTAITEM", "Alta de un Nuevo Item de Stock") < 2 Then
     Exit Sub
   End If
   '
   Command21.Enabled = False
   Call GRADATITEM
   '
   MODOAL% = 1: MODOFAM% = 0: MODORUB% = 0
   CANMODOS% = 1
   TEXMODOS$ = "Modalidad de Alta de Articulo\\Directa"
   If ULTREG&("TAFAMIL") > 0 Then
      CANMODOS% = CANMODOS% + 1
      TEXMODOS$ = TEXMODOS$ + "\Por Familia"
      MODOFAM% = CANMODOS%
   End If
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "ACHERNAR") Then
      CANMODOS% = CANMODOS% + 1
      TEXMODOS$ = TEXMODOS$ + "\Por Rubro + Subrubro"
      MODORUB% = CANMODOS%
   End If
   '
   If CANMODOS% > 1 Then
      MODOAL% = COMALTER%(TEXMODOS$)
      If MODOAL% < 1 Then GoTo 99
   End If
   '
   If MODOAL% = MODOFAM% Then ' POR FAMILIA
       ALPORFAM07.Show 1
       GoTo 90
     ElseIf MODOAL% = MODORUB% Then ' POR RUBRO
       FORPAPEL.Show 1
       GoTo 90             ' OJO QUE TAMBIEN TIENE QUE GENERAR LISTA DE SELECCION
     Else
       ' ALTA DIRECTA
       VDEF$ = String$(128, 0)
       Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
       Call REPLA(VDEF$, "U.", 109, 2)
       Call REPLA(VDEF$, Chr$(1), 101, 1)
       '
       OKX% = 0
       Call ALTACODIGO(VDEF$, OKX%)
       LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
          Text1.TEXT = CODINUE$
          SSTab1.Tab = 0
       End If
   End If
   '
90 Call ACTECOMAS
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   If DERACCE%("ALTAITEM", "Alta de un Nuevo Item de Stock") < 2 Then
     Exit Sub
   End If
   '
   Dim FIELDXVALO(256)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15

   CODIVIE$ = AJUSTI$(Text1.TEXT, LONCODI%)
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
   Call REPLA(VDEF$, MKI$(GRUCOVEI%(CII%)), 3, 2)
   Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
   Call REPLA(VDEF$, UNIMED$(CII%), 109, 2)
   Call REPLA(VDEF$, Chr$(TIMATE%(CII%)), 101, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
     Dim RMaster As ADODB.Recordset
     Set RMaster = New ADODB.Recordset
     strsql$ = "Select * from Master WITH(NOLOCK) where codint=" & CII%
     Set RMaster = READSET(strsql$)
     If Not (RMaster.EOF And RMaster.BOF) Then
        CACAM% = RMaster.Fields.Count
        For KU% = 1 To CACAM%
          DATO$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
          NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
          If NOCAMPO$ <> "CODIGO" Then
             If NOCAMPO$ <> "DESCRIPCION" Then
                If NOCAMPO$ <> "CODINT" Then
                   If NOCAMPO$ <> "ALIAS" Then
                      If NOCAMPO$ <> "UMEDIDA" Then
                        CONDI$ = "CODINT=" & CINUE%
                        Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, DATO$, REGAF&, "NOMESS")
                      End If
                   End If
                End If
             End If
           End If
        Next KU%
     End If
     '
     RMaster.Close
     Set RMaster = Nothing

'   If OKX% = 1 Then
'     Dim RMaster As ADODB.Recordset
'     Set RMaster = New ADODB.Recordset
'     RMaster.CursorType = adOpenDynamic
'     RMaster.LockType = adLockPessimistic
'     STRSQL = ("Select * from Master where codint=") & CII%
'     RMaster.Open FILTSQL$(STRSQL), FLEXCONN, , , adCmdText
'     If Not (RMaster.EOF And RMaster.BOF) Then
'        CACAM% = RMaster.Fields.Count
'        For KU% = 1 To CACAM%
'          FIELDXVALO(KU%) = RMaster.Fields(KU% - 1).Value
'        Next KU%
'        '
'        RMaster.Close
'        Set RMaster = Nothing
'        STRSQL = ("Select * from Master where codint=") & CINUE%
'        Set RMaster = New ADODB.Recordset
'        RMaster.CursorType = adOpenDynamic
'        RMaster.LockType = adLockPessimistic
'        RMaster.Open FILTSQL$(STRSQL), FLEXCONN, , , adCmdText
'        If Not (RMaster.EOF And RMaster.BOF) Then
'           For KU% = 1 To CACAM%
'             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
'             If NOCAMPO$ <> "CODIGO" Then
'               If NOCAMPO$ <> "DESCRIPCION" Then
'                 If NOCAMPO$ <> "CODINT" Then
'                   If NOCAMPO$ <> "ALIAS" Then
'                     If NOCAMPO$ <> "UMEDIDA" Then
'                       RMaster.Fields(KU% - 1).Value = FIELDXVALO(KU%)
'                     End If
'                   End If
'                 End If
'               End If
'             End If
'           Next KU%
'           RMaster.Update
'        End If
'     End If
'     '
'     RMaster.Close
'     Set RMaster = Nothing
     '
     Call COPYFORMU(CII%, CINUE%)
     Call COPYRUTA(CII%, CINUE%)
     '
     CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
     Text1.TEXT = CODINUE$
     SSTab1.Tab = 0
     UPDATMASTER% = 1
     Call GENECOMAS
     '
   End If
   '
99 Call ACTECOMAS
   Command22.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   If DERACCE%("SUSTICOD", "Sustitución de Código de Item de Stock") < 2 Then
     Exit Sub
   End If
   '
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
   
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15

   CODIVIE$ = AJUSTI$(Text1.TEXT, LONCODI%)
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > LONCODI% Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que " & LONCODI% & " Caracteres.")
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
     SQLX$ = "Update Master SET Codigo = '" & CODINUE$ & "' WHERE Codint = " & Abs(CI%)
     FLEXCONN.Execute SQLX$
     UPDATMASTER% = 1
     Call GENECOMAS
     RECODI$ = CODEXT$(CI%)
   End If
   '
   Text1.TEXT = CODINUE$
   '
99 Call ACTECOMAS
   Screen.MousePointer = 1
   Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   If DERACCE%("BAJAITEM", "Baja y Restauración de un Item de Stock") < 2 Then
     Exit Sub
   End If
   '
   Command24.Enabled = False
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), LONCODI%)
   CONDI$ = "CodiElem = " & CII%
   QTR1& = CantRegistros&("Formulas", CONDI$)
   QTR2& = CantRegistros&("FormuPAT", CONDI$)
   If QTR1& > 0 Or QTR2& > 0 Then
      'Call MENSERR(24, "Imposible Baja de este Código\por Ser Componente de Fórmula Activa.\   \Consulte Implosion y Desconecte\de los Productos Padre Correspondientes.")
      'SE INCORPORA EN EL MENSAJE LA LISTA DE CODIGOS PADRES (ACTIVOS O DE BAJA)
      SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODIELEM=" & CII%
      Set RS = READSET(SQLX$)
      II& = 0
      Do While Not RS.EOF
            CODIGO$ = CODEXT$(Abs(XVALO(RS!CODICONJ)))
            LISTACODIGOS$ = LISTACODIGOS$ & CODIGO$ & ";"
            II& = II& + 1
            If II& = 4 Then
                LISTACODIGOS$ = LISTACODIGOS$ & "\"
                II& = 0
            End If
        RS.MoveNext
      Loop
      'SE CAMBIA MENSAJE PORQUE SE HA DEJADO LA POSIBILIDAD DE ELIMINAR EL ITEM
      Call MENSERR(24, "El Codigo que Intenta dar de Baja\Es Componente de la Fórmula de Producto (Activo o de Baja)\De los Productos Padre Correspondientes:\ \" & LISTACODIGOS$)
      'GoTo 99
   End If
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   CII% = CODINT%(Text1.TEXT)
   If CII% > 0 Then
     strsql = "Update Master set Status=-1 where Codint=" & CI%
     FLEXCONN.Execute strsql, cdadreg
     UPDATMASTER% = 1
     Call BLANFORMU
     Text1 = ""
   End If
   '
99 Call GENECOMAS
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
   If Frame3.Visible = True Then
     Call MENSERR(24, "Transacción no Válida.\Cierre Cuadro 'Datos de Componente Activo'")
     Exit Sub
   End If
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
   CICONJ% = CODINT%(Text1) '
   If CantRegistros("FORMUPAT", "CodIConj = " & CICONJ%) > 0 Then
      Call MENSERR(24, "Debe Actualizar Unicamente\por Fórmula Patrón.")
      Exit Sub
   End If
   '
  CIELEM% = CODINT%(Text31)
   MAXCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   
   For IX% = 0 To List1.ListCount - 1
     CODIGO$ = TRIM$(Mid(List1.List(IX%), 1, MAXCODI%))
     If CODIGO$ = TRIM$(Text31) Then
        RTA1% = COMALTER%("Codigo " & CODIGO$ & " Pre-Existente en Fórmula\¿Realmente Desea Agregarlo Repetido?\\No, Cancelar\Si, Repetir")
        If RTA1% <> 2 Then Exit Sub
        Exit For
     End If
   Next IX%
 
   '--Fin Validacion Item en formula------------*
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = XVALO(Text32.TEXT)
   REFCO$ = TRIM$(Text39)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(XVALO(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(128)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
   '\\\OT-05-0633-WAR-23/09/05///
   If Frame16.Visible = False Then
     Text41 = TRIM$(USOX$)
   End If
   Call REPLA(XXX$, Text41.TEXT, 89, 9)
   Call REPLA(XXX$, Label42, 101, 7)
   Call REPLA(XXX$, Text42.TEXT, 111, 6)
   Call REPLA(XXX$, Label36, 119, 8)
   '\\\OT-05-0633-WAR-FIN///
   XXX$ = XXX$ + REFCO$
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
   Text39 = ""
   Text41 = ""
   Text42 = ""
   Text31.SetFocus
   MODIFIC%(2) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
   If Frame3.Visible = True Then
     Call MENSERR(24, "Transacción no Válida.\Cierre Cuadro 'Datos de Componente Activo'")
     Exit Sub
   End If
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
   CICONJ% = CODINT%(Text1)
   If CantRegistros("FORMUPAT", "CodIConj = " & CICONJ%) > 0 Then
      Call MENSERR(24, "Debe Actualizar Unicamente\por Fórmula Patrón.")
      Exit Sub
   End If
   '
   '-- Valida Si Item existe en la formula-------*
   CIELEM% = CODINT%(Text31)
   MAXCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   
   For IX% = 0 To List1.ListCount - 1
     CODIGO$ = TRIM$(Mid(List1.List(IX%), 1, MAXCODI%))
     If CODIGO$ = TRIM$(Text31) Then
        RTA1% = COMALTER%("Codigo " & CODIGO$ & " Pre-Existente en Fórmula\¿Realmente Desea Agregarlo Repetido?\\No, Cancelar\Si, Repetir")
        If RTA1% <> 2 Then Exit Sub
        Exit For
     End If
   Next IX%
 
   '--Fin Validacion Item en formula------------*
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = XVALO(Text32.TEXT)
   REFCO$ = TRIM$(Text39)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(XVALO(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(128)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
    '\\\OT-05-0633-WAR-23/09/05///
   If Frame16.Visible = False Then
     Text41 = TRIM$(USOX$)
   End If
   Call REPLA(XXX$, Text41.TEXT, 89, 9)
   Call REPLA(XXX$, Label42, 101, 7)
   Call REPLA(XXX$, Text42, 111, 6)
   Call REPLA(XXX$, Label36, 119, 8)
   '\\\OT-05-0633-WAR-FIN///
   XXX$ = XXX$ + REFCO$
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
   Text39 = ""
   Text41 = ""
   Text42 = ""
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

Private Sub Command29_Click(Index As Integer)
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33(Index).TEXT = VALONUE$
   End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If MODOEDIT% = 1 Then Call MENSAEDIT: GoTo 99
   Call GRADATITEM
   Call CIERRARCH("*.*")
   If App.EXEName = "AMASTO07" Then
        Call FINAL("")
      Else
        Unload Me
   End If
99 Command3.Enabled = True
End Sub

Private Sub Command30_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command30.Enabled = False
   If EXISTMU% = 0 Then '\\\OT-06-0252-WAR-26/06/06///
     Call GRADATITEM
     For KKII% = 0 To 15: YACAR%(KKII%) = 0: Next KKII%
     Call COPYRUTAS
   ElseIf EXISTMU% = 1 Then
     RTA1% = COMALTER%("Seleccione la Opción de Copia:\\Operaciones Std.\Operaciones TMU's")
     If RTA1% = 1 Then
       Call GRADATITEM
       For KKII% = 0 To 15: YACAR%(KKII%) = 0: Next KKII%
       Call COPYRUTAS
     ElseIf RTA1% = 2 Then
       FOCOPOP07.Text10.TEXT = Text1.TEXT
       FOCOPOP07.Show 1
     End If
   End If
   Command30.Enabled = True
End Sub

Private Sub COMMAND31_CLICK()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
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
   Label52.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 5
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30(0).TEXT = ""
   Text30(1).TEXT = ""
   Text33(0).TEXT = ""
   Text33(1).TEXT = ""
   Text34.TEXT = ""
   Label51(0).Caption = ""
   Label51(1).Caption = ""
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
     NOPY% = 10 * (1 + Int(XVALO(Left$(List3.List(ULI& - 1), 4) / 10)))
   End If
   If List3.ListCount >= MAXECO% Then
      Call MENSERR(24, "Imposible Agregar")
      Exit Sub
   End If
   RESECO$(List3.ListCount) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(128, 0) 'para acomodar hoja de procesos DE FABRICACION
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

Sub Command37_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   For KKI% = 0 To 15: YACAR%(KKI%) = 0: Next KKI%
   '
   If DERACCE%("STPMAST", "Configuracion de Setup de Articulos") = 2 Then
     OPMAST07.Show 1
     Call FINAL("")
   End If
   '
End Sub

Sub Command38_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command38.Enabled = False
   OPLIMAS07.Show 1
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
    Command39.Enabled = False
    Dim Nuevo_Form As Form
      
    Set Nuevo_Form = New AMASTOK07
      
    UTILIZA_SKIN% = 0
    With Nuevo_Form
        .PRINTFORM
        .Hide
    End With
    '
    UTILIZA_SKIN% = 1
99 Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call HELPONLINE("AMASTO07")
   Command4.Enabled = True
End Sub

 Sub Command40_Click(Index As Integer)
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

Sub Command42_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command42.Enabled = False
   If EXISTMU% = 0 Then '\\\OT-06-0252-WAR-26/06/06///
     Call GRADATITEM
     Call LISTRUTAS
   ElseIf EXISTMU% = 1 Then
     RTA1% = COMALTER%("Seleccione la Opción de Listado:\\Operaciones Std.\Operaciones TMU's")
     If RTA1% = 1 Then
       Call GRADATITEM
       Call LISTRUTAS
     ElseIf RTA1% = 2 Then
       Call LISTMUS
     End If
   End If
   Command42.Enabled = True
End Sub

Private Sub Command45_Click()
   '
   Frame3.Visible = False
   '
End Sub

'Private Sub Command48_Click()
'   COMIE = Text43.SelStart
'   LOSEL = Text43.SelLength
'   Text43.TEXT = Left$(Text43.TEXT, COMIE) + Clipboard.GetText + Mid$(Text43.TEXT, COMIE + LOSEL + 1)
'End Sub

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
   NOPEP% = XVALO(Left$(List3.List(ILISTAX&), 4))
   NOPEA% = 0
   If ILISTAX& > 0 Then
     NOPEA% = XVALO(Left$(List3.List(ILISTAX& - 1), 4))
   End If
   NOPY% = (NOPEA% + NOPEP%) / 2
   If NOPY% <= NOPEA% Or NOPY% >= NOPEP% Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   If List3.ListCount >= MAXECO% Then
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
   RESECO$(ILISTAX&) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(128, 0) 'para acomodar hoja de procesos DE FABRICACION
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


Private Sub Command9_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   FORSELEC.ABUSQUE = Text1
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Frame3_DblClick()
   If Command8.Enabled = True Then
     Call Command18_Click
   End If
End Sub

Private Sub guar_Click()
   Call Command1_Click
End Sub

Private Sub Image1_Click()
   Call Label47_DblClick(0)
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

Private Sub Label10_DblClick(Index As Integer)
    If EXISTE%("TRAZAB07.EXE") > 0 Then
       ACOSLOTE.Label28 = Text1
       ACOSLOTE.Show 1
    End If
End Sub

Private Sub Label50_DblClick()
   Text47.Visible = True
   Text47 = Label50
   Text47.SetFocus
End Sub

Private Sub Label61_DblClick()
     Call label58_Click
End Sub

Private Sub Label63_Change()
     MODIFIC%(1) = 1
End Sub

Private Sub Label63_DblClick()
    If TRIM$(Label63) <> "" Then
       If (EXISTE%(Label63)) <> 1 Then
          Call COMUNI("Archivo Inexistente")
          Exit Sub
       End If
       Call MUESTRAPDF(TRIM$(Label63))
    End If
    
End Sub

Private Sub Label71_Change()
   '
   UVALUE% = XVALO(Label71)
   If UVALUE% = 0 Then Label18 = "Susp.Sin Movim."
   If UVALUE% = 1 Then Label18 = "Activo"
   If UVALUE% = 2 Then Label18 = "Autorizado"
   '

End Sub

Private Sub Label71_Click()
    If DERACCE%("MODISTAT", "Modificación Estado de Movimientos de Artículos") >= 2 Then
        MODIFIC%(0) = 1
        '
        UVALUE% = XVALO(Label71)
        UVALUE% = (UVALUE% + 1) Mod 3
        Label71 = FORMATNUM$(UVALUE%, "I1")
    End If
End Sub

Private Sub List2_DblClick()
   LINEA$ = List2.List(List2.ListIndex)
   Archi$ = TRIM$(Mid$(LINEA$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If
End Sub



Private Sub Picture5_DblClick()
    Call Label27_DblClick
End Sub

Private Sub Text11_DblClick()
16 TTTT$ = TRIM$(Text11)
   If TTTT$ <> "" Then
      If EXISTE%(LUDAT$ + TTTT$ + ".SQR") > 0 Then
         ORDHOJA& = 1 + List3.ListIndex
         TOTHOJA& = List3.ListCount
         '
         FILTX$ = "Id-Lote" & ";" & Label28 & ";" & Label55
         Call STDFORM(TTTT$, FILTX$)
         STDSCREEN.Caption = Text11
         Openforms = DoEvents
         If STDSCREEN.PAGIN = "NEXT" Then          ' PAGINA SIGUIENTE
            If ORDHOJA& < List3.ListCount Then
               List3.ListIndex = ORDHOJA&
               GoTo 16
            End If
          ElseIf STDSCREEN.PAGIN = "PREV" Then          ' PAGINA ANTERIOR
            If ORDHOJA& > 1 Then
               List3.ListIndex = ORDHOJA& - 2
            End If
            GoTo 16
          Else
            GoTo 99
         End If
       ElseIf EXISTE%(LUDAT$ + TTTT$ + ".PDF") > 0 Then
         HOPROFA04.Pdf1.setShowToolbar (True)
         On Error Resume Next
         HOPROFA04.Pdf1.LOADFILE LUDAT$ + TTTT$ + ".PDF"
         If Err Then GoTo 20
         On Error GoTo 0
         HOPROFA04.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo " + TTTT$)
   '
99 End Sub


   '


'\\\OT-05-0720-WAR-25/11/05///
Private Sub Text40_LostFocus()
   Text40.TEXT = CSTRING$(MKS$(XVALO(Text40.TEXT)), 3, 10, 1, 2)
End Sub
Private Sub Text49_LostFocus() ' text del lote indivisible de ventas
   Text49.TEXT = CSTRING$(MKS$(XVALO(Text49.TEXT)), 3, 10, 1, 2)
End Sub

'\\\OT-05-0720-WAR-FIN///
'\\\OT-05-0633-WAR-22/09/05///
Sub TEXT41_CHANGE()
  '
  CODEX1$ = Text31
  CI1% = CODINT(CODEX1$)
  Call CALCUKILOS(CI1%)
  '
End Sub
Sub TEXT42_CHANGE()
  '
  CODEX1$ = Text31
  CI1% = CODINT(CODEX1$)
  Call CALCUKILOS(CI1%)
  '
End Sub
Sub LABEL42_CHANGE()
  '
  CODEX1$ = Text31
  CI1% = CODINT(CODEX1$)
  Call CALCUKILOS(CI1%)
  '
End Sub
Sub TEXT41_LOSTFOCUS()
  '
  VALO1 = XVALO(Text41)
  TF$ = FORMATNUM$(VALO1, "F9.4")
  Text41 = TF$
  '
End Sub

Private Sub Text42_GotFocus()
   Text42.SelStart = 0
   Text42.SelLength = Len(Text42.TEXT)
End Sub

Sub TEXT42_LOSTFOCUS()
  '
  VALO1 = XVALO(Text42)
  TF$ = FORMATNUM$(VALO1, "F6.1")
  Text42 = TF$
  '
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

Private Sub Label19_DblClick()
   Frame3.Visible = False
   Call Text31_DblClick
End Sub

Private Sub Label46_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 159, 2) = MKI(XVALO(Mid$(Label46, 49, 6)))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = XVALO(Mid$(Label46, 49, 6))
   TERPESOS = XVALO(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   MODIFIC%(3) = 1
End Sub

Private Sub Label46_DblClick()
   '
   Call SELECHO("TAMONED")
   MONEI% = XVALO(TRIM$(VALACT1$("TAMONED")))
   If MONEI% > 0 Then
     MONEX$ = AJUSTI$(VALACT2$("TAMONED"), 48) + Str$(MONEI%)
     Label46.Caption = MONEX$
   End If
   '
End Sub

Private Sub Label47_DblClick(Index As Integer)
   If Index = 0 Then
     If XVALO(Label55) > 0 Then
       VDEF$ = Space$(64) + String$(64, 0)
       Call REPLA(VDEF$, Text1, 1, 16)
       Call REPLA(VDEF$, Text3, 17, 48)
       Call REPLA(VDEF$, MKI$(XVALO(Label55)), 65, 2)
       Call REPLA(VDEF$, Text5, 67, 30)
       '
       CIXI% = CODINT%(Text1): NOPYI% = XVALO(Label55)
       TPXX = XVALO(Text15(0)) + XVALO(Text17(0)) / 60 + XVALO(Text26(0)) / 3600
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
         SQLX$ = " SELECT * FROM SECOPER"
         SQLX$ = SQLX$ + " WHERE CodIConj = " & CIXI% & " "
         SQLX$ = SQLX$ + " AND NumeOper = " & NOPYI% & " "
         SQLX$ = SQLX$ + " AND StatOper > 0 "
         Dim SECCCOPER As ADODB.Recordset
         Set SECCCOPER = New ADODB.Recordset
         SECCCOPER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

         With SECCCOPER
           '.FindFirst ("CodIConj = " & CIXI% & "AND NumeOper = " & NOPYI% & " AND StatOper > 0")
           'If .NoMatch = False Then
               '.Edit
           If Not (.EOF And BOF) Then
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
         'Call CALCOSOPE
         Screen.MousePointer = 1
         '
       End If
     End If
   End If
   '
   'NUEVO PARA MANTENIMIENTO HERRAMENTAL
   If Index = 4 Then
     If XVALO(Text1) > 0 Then
       NOPE% = XVALO(List3.TEXT)
       If NOPE% > 0 Then
         CLAVE$ = MKI$(CI%) + MKI$(NOPE%)
         Call FILTERFILE("MANTHERR", "", 1, CLAVE$)
         VTB% = VENTABU%("MANTHERR")
         If VTB% >= 0 Then
           TOTMINU = 0
           For II& = 1 To ULTREG("!MANTHERR")
             x$ = REGLEIDO("!MANTHERR", II&)
             MINUL = CVS(Mid$(x$, 33, 4))
             CAPIE = CVS(Mid$(x$, 37, 4))
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
   'Nuevo para TMU's
   '\\\OT-06-0252-WAR-22/06/06///
   If EXISTMU% = 1 Then
    If Index = 2 Then
      If XVALO(Label55) > 0 Then 'Si Tiene Operacion Seleccionada
        CIXI% = CODINT%(Text1): NOPYI% = XVALO(Label55)
        Call ABRESECOP
        Call ABREMASTER
        '
        SQLX$ = "SELECT * FROM TMUSOPER "
        SQLX$ = SQLX$ + " WHERE Cod_Inte= " & CIXI% & " "
        SQLX$ = SQLX$ + " AND Cod_Oper = " & NOPYI% & ""
        SQLX$ = SQLX$ + " ORDER BY Nume_Ord "
        'Set TMUTMP = Articulos.OpenRecordset(SQLX$, 4)
        Dim TMUTMP As ADODB.Recordset
        Set TMUTMP = New ADODB.Recordset
        TMUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

        With TMUTMP
          JJ& = 0
          REBLA1$ = REGBLAN$("TACATMU")
          On Error Resume Next
          .MoveFirst
          If Err Then
            On Error GoTo 0
            GoTo 10
          End If
          On Error GoTo 0
          Do While .EOF = False
            ZZ$ = REBLA1$
            CANTIX# = !Cantidad
            COTMU$ = !Cod_TMU
            MSEG2# = MSTMU#(COTMU$) 'ME DEVUELVE LOS MILISEGUNDOS DEL TMU
            '
            Call REPLA(ZZ$, COTMU$, 1, 4)
            Call REPLA(ZZ$, MKD$(CANTIX#), 13, 8)
            Call REPLA(ZZ$, MKD$(MSEG2#), 5, 8)
            JJ& = JJ& + 1
            Call GRAREG("!TACATMU", ZZ$, JJ&)
            .MoveNext
          Loop
        End With
        TMUTMP.Close
        Set TMUTMP = Nothing
        '
10      Call SETULTREG("!TACATMU", JJ&)
        Call FRESHECHO("!TACATMU")
        ATRI1$ = "Pieza: " + Text1.TEXT
        ATRI1$ = ATRI1$ + " - Operación: " + Text5.TEXT
        VTB% = VENTABU%("TACATMU/TITUPAN=" + ATRI1$)
        If VTB% >= 0 Then
'          With TmusOper 'BORRO TODOS
'            SQLX$ = "Cod_Inte= " & CIXI% & " AND Cod_Oper = " & NOPYI% & ""
'            .FindFirst SQLX$
'15          If .NoMatch = False Then
'             .Delete
'             .FindNext SQLX$
'             GoTo 15
'            End If
'          End With
           SQLX$ = "DELETE FROM  TmusOper"
           SQLX$ = SQLX$ + " WHERE Cod_Inte= " & CIXI% & " "
           SQLX$ = SQLX$ + " AND Cod_Oper = " & NOPYI% & ""
            FLEXCONN.Execute (FILTSQL$(SQLX$))
          '
          TISTOP = 0
          SQLX$ = "SELECT * FROM TmusOper "
          Dim TMUSSSOPER As ADODB.Recordset
          Set TMUSSSOPER = New ADODB.Recordset
          TMUSSSOPER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

          For II& = 1 To ULTREG("!TACATMU")
            x$ = REGLEIDO("!TACATMU", II&)
            COTMU$ = Mid$(x$, 1, 4)
            CANTI1 = CVD(Mid$(x$, 13, 8))
            MSEG1# = MSTMU#(COTMU$) 'ME DEVUELVE LOS MILISEGUNDOS DEL TMU
            TISTOP = TISTOP + (CANTI1 * MSEG1# / 1000)
            '
            With TMUSSSOPER
              .AddNew 'CARGO TODOS   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
                !COD_INTE = CIXI%
                !Cod_Oper = NOPYI%
                !Cod_TMU = COTMU$
                !Nume_Ord = II&
                !Cantidad = CANTI1
              .Update
            End With
            '
          Next II&
          Call FRESHECHO("!TACATMU")
          '
          CABAS = XVALO(Text7.TEXT)
          TISTA = CABAS * TISTOP
          TSTT$ = HORMINSEDIEM$(TISTA)
          Text15(2).TEXT = Left$(TSTT$, 3)
          Text17(2).TEXT = Mid$(TSTT$, 5, 2)
          Text26(2).TEXT = Mid$(TSTT$, 8, 7)
          '
        End If
      End If
    End If
  End If
   '\\\OT-06-0252-WAR-FIN///
End Sub

Private Sub LisPre_Cli()
   Call CONECRUN("LISPRE07", "Lista de Precios de Venta")
End Sub

Private Sub List1_Click()
    If MODOEDIT% = 0 Then
      Frame3.Visible = True
      Label19 = Left$(List1.TEXT, 16)
      CDYY% = CODINT%(Label19)
      Label4 = DESCRIT0$(CDYY%)
      Label9 = UNIMED$(CDYY%)
      '\\\OT-05-0633-WAR-23/09/05///
      Label24 = TRIM$(Mid$(List1.TEXT, 89, 9))
      Label25 = TRIM$(Mid$(List1.TEXT, 101, 7))
      Label26 = TRIM$(Mid$(List1.TEXT, 111, 6))
      Label34 = TRIM$(Mid$(List1.TEXT, 119, 8))
      '\\\OT-05-0633-WAR-FIN///
      Label25 = FORMATNUM$(PESMETRO(CDYY%), "F7.3")
      Label20 = TRIM$(Mid$(List1.TEXT, 49, 12))
      Label22 = TRIM$(Mid$(List1.TEXT, 61, 4))
      Label23 = TRIM$(Mid$(List1.TEXT, 129))
      Check2.Value = PORMILLAR%
    End If
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
   '\\\OT-06-0252-WAR-29/06/06///
   If EXISTMU% = 1 Then 'Si Existe Tabla de TMU's
     Call ABREMASTER
     Call ABRESECOP
     CIXI% = CODINT%(Text1)
     NOPYI% = XVALO(Mid$(List3.List(List3.ListIndex), 1, 4))
     SQLX$ = "SELECT * FROM TmusOper "
     SQLX$ = SQLX$ + "WHERE Cod_Inte=" & CIXI% & " AND Cod_Oper = " & NOPYI% & ""
     Dim TMUSSSOPER As ADODB.Recordset
     Set TMUSSSOPER = FLEXCONN.Execute(FILTSQL$(SQLX$))

     'With TmusOper
       'SQLX$ = "Cod_Inte= " & CIXI% & " AND Cod_Oper = " & NOPYI% & ""
       '.FindFirst SQLX$
       'If .NoMatch = False Then
     With TMUSSSOPER
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Text15(2).Enabled = False 'No se pueden editar el tiempo variable
         Text17(2).Enabled = False 'surge por calculo de TMU's
         Text26(2).Enabled = False
       Else
        On Error GoTo 0
        Text15(2).Enabled = True 'Se pueden editar el tiempo variable
        Text17(2).Enabled = True
        Text26(2).Enabled = True
       End If
     End With
     TMUSSSOPER.Close
     Set TMUSSSOPER = Nothing
   End If
   '\\\OT-06-0252-WAR-FIN///
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
   
   Text51.TEXT = ""
   Text52.TEXT = ""
   Text3.TEXT = ""
   Text38.TEXT = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
   Text4.TEXT = ""
   Label30.Caption = ""
   Text6.TEXT = ""
   Label29.Caption = ""
   Text48.TEXT = ""
   Label64.Caption = ""
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
   Text44 = ""
   Text46 = ""
   Text49 = ""
   Label39.Caption = ""
   Label67.Caption = ""
   Label69.Caption = ""  'COSTO CALCULADO
   Label70.Caption = "" 'Fecha del calculo
   '
   List1.Clear
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   ' limpia el frame3
   Label19 = ""
   Label4 = ""
   Label9 = ""
   Label20 = ""
   Label22 = ""
   Label23 = ""
   '\\\OT-05-0633-WAR-23/09/05///
   Label24 = ""
   Label25 = ""
   Label26 = ""
   Label34 = ""
   '\\\OT-05-0633-WAR-FIN///
   '
   List3.Clear
   Label55.Caption = ""
   Label52.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 5
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30(0).TEXT = ""
   Text30(1).TEXT = ""
   Text33(0).TEXT = ""
   Text33(1).TEXT = ""
   Label51(0).Caption = ""
   Label51(1).Caption = ""
   Text45 = ""
   Text43 = "" ' TEXT DE ESPECIFICACION
   Text81 = ""
   Text82 = ""
   Text83 = ""
   Text84 = ""
   '
   Check3.Value = False  '\\\OT-06-0105-WAR-23/03/06///
   Label71 = ""
   Label58 = ""
   Label63 = "" ' MEMORIA  TECNICA
   '
   For KKII% = 0 To 15
     YACAR%(KKII%) = 0
     MODIFIC%(KKII%) = 0
   Next KKII%

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
    Text1.TEXT = CODEXT$(CI%)
    Text3.TEXT = DESCRIT0$(CI%)
    Text38.TEXT = DESCRIT2$(CI%)
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
        Call CARESPECI
    End Select
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
         FAMI$ = CODFAMIL$(CI%)
      Text48.TEXT = FAMI$
      Label64.Caption = ECOARCH$("TAFAMIL", FAMI$)
         GRUCO% = GRUCOVEI%(CI%)
      Text16.TEXT = TRIM$(Str$(GRUCO%))
      Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(GRUCO%))
         GRUIVA% = GRUPIVA%(CI%)
      Text35.TEXT = TRIM$(Str$(GRUIVA%))
      Label57.Caption = ECOARCH$("GRUPIVA", Chr$(GRUIVA%))
      '
      ' SI TIENE DERECHO DE VER EL COSTO - SI NO PRESENTA ASTERIZCOS
      'PUEDE VER PERO LA CAJA DE TEXTO SE DESHABLITA
      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 0 Then
        Text12.TEXT = CSTRING$(MKS$(COSUNI(CI%)), 3, 11, 4, 2)
        Text12.Enabled = False
        'SI TIENE DERECHOS DE MODIFICAR COSTOS SE HABILITA LA CAJA DE TEXTO
        If DERACCE%("MODCOSTO/NOMESS", "Modificar Costos y Precios") > 1 Then
           Text12.Enabled = True
        End If
      Else
        'PRESENTA ASTERIZCOS PARA QUE NO SE VISUALICE EL PRECIO
        Text12 = FORMATNUM$(99999999999#, "F11.4")
        Text12.Enabled = False
      End If

         FEX = FECOSTO%(CI%)
      Text14.TEXT = FECHATEX$(FEX)
         MONEI% = MONECOSTO%(CI%)
      Text13.TEXT = TRIM$(Str$(MONEI%))
      Label35.Caption = ECOARCH$("TAMONED", Chr$(MONEI%))
         PORNA = PORNACIO(CI%)
      Text28.TEXT = AJUSTD$(FORMATNUM$(PORNA, "F7.2"), 8)
      '
      COSUDOL = XVALO(Text12)
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
      '
      '\\\OT-05-0720-WAR-25/11/05///
      On Error Resume Next
      Text40.TEXT = CSTRING$(MKS$(LINDIVIS(CI%)), 3, 10, 1, 2)
      Text49.TEXT = CSTRING$(MKS$(LINDIVEN(CI%)), 3, 10, 1, 2)
      On Error GoTo 0
      '\\\OT-05-0720-WAR-FIN///
      '
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(CI%)), 3, 10, 1, 2)
      '
      CONDI$ = "Codint = " & CI%
      Text22.TEXT = COLOR$(CI%)
      Text23.TEXT = TALLE$(CI%)
      Text44.TEXT = VALOBADA("MASTER", "COD_INTERNA", CONDI$, "NOMESS")
      Text24.TEXT = LTITPE$(PESUNI(CI%))
      '
      '\\\OT-5-0633-FG-21/09/05
      Text46.TEXT = FORMATNUM$(PESMETRO(CI%), "F8.3")
      '\\\OT-5-0633-FG-21/09/05
      '
      Text36 = REVICOD$(CI%)
      Text25 = FECHATEX$(FEREVI%(CI%))
      '
      '\\\OT-06-0105-WAR-23/03/06///
      If DESCOMPO%(CI%) = 1 Then Check3.Value = 1
      '\\\OT-06-0105-WAR-FIN///
      Label58 = FORMATNUM$(ESTRAZABLE%(CI%), "I1")
      '
      Text50.TEXT = SAFETEXT(VALOBADA("MASTER", "CANDECI", CONDI$, "NOMESS"))
      CONS# = XVALO(VALOBADA("MASTER", "consuprom", CONDI$))
      Label67 = FORMATNUM(CONS#, "F13.1")
      '
      COSCAL# = XVALO(VALOBADA("MASTER", "CosCal", CONDI$)) 'Costo Calculado
      Label69 = FORMATNUM(COSCAL#, "F13.4")
      '
      Label70 = FECHATEX(CVINT(VALOBADA("MASTER", "FECOSCAL", CONDI$)))
      '
      Label71 = "1"
      If STATMOVI%(CI%) = -1 Then Label71 = ""
      If STATMOVI%(CI%) = 1 Then Label71 = "2"
      '
      'Se modificaron estas dos lecturas para que traigan solo el dato de la base
      'Muestra el Valor del Lead_Time ingresado
      
      Text51.TEXT = VALOBADA("MASTER", "LeTiCom", CONDI$, "NOMESS")
      '
      'PERIODO DE STOCKS MINIMO
      'no se lee con xvalo para que no  traiga 0 en el caso que este vacio
      Text52.TEXT = VALOBADA("MASTER", "Autono_Proy", CONDI$, "NOMESS")
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
      Text43.TEXT = TRIM$(ESPECIFIT$(CI%))
      Label63.Caption = ARMEMOTEC$(CI%)
      '
      Text81 = TRIM$(VALOBADA("Master", "Clas_ATC", "codint = " & CI%, "NOMESS"))
      Text82 = TRIM$(VALOBADA("Master", "Vers_BC", "codint = " & CI%, "NOMESS"))
      Text83 = TRIM$(FECHATEX$(CVINT(VALOBADA("Master", "FeVig_BC", "codint = " & CI%, "NOMESS"))))
      Text84 = TRIM$(FECHATEX$(CVINT(VALOBADA("Master", "FeRev_BC", "codint = " & CI%, "NOMESS"))))
      YACAR%(1) = 1
      '
      List2.Clear
      SQLX$ = "SELECT * FROM DOCUREL WITH(NOLOCK) WHERE COD_INTE = " & CI%
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           rutax$ = SAFETEXT$(!RUTA$)
           Call REPLA(TX$, DESCRI$, 1, 24)
           Call REPLA(TX$, rutax$, 32, 512)
           List2.AddItem TX$
          .MoveNext
        Loop
      End With
      
      MODIFIC%(1) = 0
      '
   End If
   '
End Sub
'
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
      PORMILLAR% = 0
      'SE AGREGO PARAMETRO 'COMPLETO' A LEEEXPLO PARA QUE CUANDO ES 1 INCLUYE TAMBIEN A LOS ANULADOS
      Call LEEEXPLO(CODD$, 1, 1)
      Label59.Caption = FECHATEX$(FEREFOR%(CI%))
      If TRIM$(Label59.Caption) <> "" Then
         CONDI$ = "CODINT = " & CI%
         Label59.ToolTipText = "Aprobada por " & SAFETEXT$(VALOBADA("Master", "UsuApFor", CONDI$, "NOMESS"))
      End If
      Label59.Refresh
      List1.Clear
      NOCHECK% = 1
      Check1.Value = 0: If PORMILLAR% = 1 Then Check1.Value = 1
      NOCHECK% = 0
'     Decide si muestra o no el Boton aprobado
      If CONTROL("FORMULAS", "REQAPRO") = "SI" Then
         Frame15.Caption = "Aprobacion"
         USU$ = SAFETEXT(VALOBADA("MASTER", "USUAPFOR", "CODINT = " & CI%))
         If USU$ = "PENDIENTE DE APROBACION" Then
              Command51.Visible = True
            Else
              Command51.Visible = False
              Label59.Caption = FECHATEX$(FEREFOR%(CI%))
              Label59.ToolTipText = "Aprobada por " & USU$
         End If
      End If
      
      '
      For I% = 1 To CAIT%
         List1.AddItem EXPLOLIN$(I%)
      Next I%
      '
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
      Label52 = ""
      Text5 = ""
      Text7 = ""
      Text11 = ""
      Label43 = ""
      Text29 = ""
      Text30(0) = ""
      Text30(1) = ""
      Text33(0) = ""
      Text33(1) = ""
      Text34 = ""
      Label51(0) = ""
      Label51(1) = ""
      For III% = 0 To 4
        Text15(III%) = ""
        Text17(III%) = ""
        Text26(III%) = ""
      Next III%
      For III% = 0 To 5
        Text9(III%) = ""
      Next III%
      '
      Label60 = ""
      Label62 = ""
      ULTOP% = 0
      SQLX$ = "SELECT * FROM Secoper"
      SQLX$ = SQLX$ & " WHERE CodIConj = " & CI%
      SQLX$ = SQLX$ & " AND StatOper >0 "
      SQLX$ = SQLX$ + "ORDER BY NumeOper, AltOper ASC"
      
      'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim SECOPTEMP As ADODB.Recordset
      Set SECOPTEMP = New ADODB.Recordset
      SECOPTEMP.CursorType = adOpenKeyset
      SECOPTEMP.LockType = adLockReadOnly
      SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
           
      '
      REBLA$ = REGBLAN$("SECOPER") + String$(256, 0)
      IOP% = 0
      On Error Resume Next
      With SECOPTEMP
        .MoveFirst
        If Err Then GoTo 25
        '
        Do Until .EOF = True
          NOPY% = !NumeOper
          ALTY% = !AltOper
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
          CAOPS2 = !CantiOps2
          CATOP% = !CategOps
          CATOP2% = !CategOps2
          EQUIPO$ = !CodMaq0
          EQUIP1$ = !CodMaq1
          EQUIP2$ = !CodMaq2
          EQUIP3$ = !CodMaq3
          EQUIP4$ = !CodMaq4
          EQUIP5$ = !CodMaq5
          '
          OBSOPX$ = !OBSEOPER
          '
          PROTERI% = !ProTerOp
          PRECTER# = !PreUnid
          MONETER% = !MonePrec
          FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
          '\\\OT-06-0046-WAR-10/02/06///
            HOPFAB1$ = !HoProFab
          '\\\OT-06-0046-WAR-FIN///
          Label60 = FECHATEX$(CVINT(!FECH_REVIS))
          Label62 = SAFETEXT$(!USUA_REVIS)
          '
          x$ = REBLA$
          Call REPLA(x$, MKI$(CI%), 1, 2)
          Call REPLA(x$, MKI$(NOPY%), 3, 2)
          Call REPLA(x$, Chr$(ALTY%), 88, 1)
          Call REPLA(x$, DOPE$, 5, 48)
          Call REPLA(x$, EQUIPO$, 53, 6)
          Call REPLA(x$, MKS$(HORPREP), 59, 4)
          Call REPLA(x$, MKS$(HORFIJO), 63, 4)
          Call REPLA(x$, MKS$(MINVARI), 67, 4)
          Call REPLA(x$, MKS$(DEMODIS), 75, 4)
          Call REPLA(x$, MKS$(HORLIMP), 79, 4)
          Call REPLA(x$, MKS$(CAOPS), 83, 4)
          Call REPLA(x$, MKS$(CAOPS2), 178, 4)
          Call REPLA(x$, Chr$(CATOP%), 87, 1)
          Call REPLA(x$, Chr$(CATOP2%), 177, 1)
          Call REPLA(x$, MKI$(SOLAPAI%), 89, 2)
          Call REPLA(x$, MKS$(LOTESTA), 91, 4)
          Call REPLA(x$, MKS$(HORHERR), 95, 4)
          Call REPLA(x$, MKS$(PIECICL), 99, 4)
          Call REPLA(x$, EQUIP1$, 103, 6)
          Call REPLA(x$, EQUIP2$, 109, 6)
          Call REPLA(x$, EQUIP3$, 115, 6)
          Call REPLA(x$, EQUIP4$, 182, 6)
          Call REPLA(x$, EQUIP5$, 188, 6)
          '
          '\\\OT-06-0046-WAR-10/02/06///
          Call REPLA(x$, HOPFAB1$, 129, 16)
          '\\\OT-06-0046-WAR-FIN///
          '
          Call REPLA(x$, MKI$(PROTERI%), 149, 2)
          Call REPLA(x$, MKD$(PRECTER#), 151, 8)
          Call REPLA(x$, MKI$(MONETER%), 159, 2)
          Call REPLA(x$, MKI$(FEPRETER%), 161, 2)
          
          If NOPY% > ULTOP% Then ULTOP% = NOPY%
          NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
          SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
          If ALTY% > 0 Then
            Mid$(SECO$, 5, 5) = "-" + TRIM$(Str$(ALTY%))
          End If
          List3.AddItem SECO$
          RESECO$(List3.ListCount - 1) = x$  ' Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0)
          DESOPE$(List3.ListCount - 1) = OBSOPX$
          '
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
   REGI& = XVALO(Mid$(List3.TEXT, 59, 6))
   NOPY% = XVALO(Left$(List3.TEXT, 4))
   ALTY% = XVALO(Mid$(List3.TEXT, 6, 2))
   DOPE$ = Mid$(List3.TEXT, 10, 48)
   '
   VDEF$ = RESECO$(List3.ListIndex)
   '
   Label55.Caption = TRIM$(Str$(NOPY%))
   Label52.Caption = TRIM$(Str$(ALTY%))
   Text5.TEXT = DOPE$
   Text7.TEXT = CSTRING(MKS$(CABAS), 3, 7, 0, 2)
   '
   Text9(0).TEXT = AJUSTI$(Mid$(VDEF$, 53, 6), 6)
   Label43.Caption = ECOARCH$("PADMAQ", Text9(0).TEXT)
   Text9(1).TEXT = Mid$(VDEF$, 103, 6)
   Text9(2).TEXT = Mid$(VDEF$, 109, 6)
   Text9(3).TEXT = Mid$(VDEF$, 115, 6)
   Text9(4).TEXT = Mid$(VDEF$, 182, 6)
   Text9(5).TEXT = Mid$(VDEF$, 188, 6)
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
   Text30(0).TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 83, 4), 3, 5, 2, 2))
   Text30(1).TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 178, 4), 3, 5, 2, 2))
   Text33(0).TEXT = TRIM$(Str$(Asc(Mid$(VDEF$, 87, 1))))
   Text33(1).TEXT = TRIM$(Str$(Asc(Mid$(VDEF$, 177, 1))))
   Text34.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 99, 4), 3, 5, 0, 2))
   Label51(0).Caption = ECOARCH$("CATOPER", Chr$(Text33(0).TEXT))
   '
   Text2.TEXT = RTrim$(DESOPE$(List3.ListIndex))
   '

   'VALIDAR SI TIENE FECHA EN SQL CASO CONTRARIO BUSCA EN DATADOPS PARA NO PERDER LA INFORMACION
   NOPER% = XVALO(Label55)
   CONDI$ = " CodIConj = " & CI% & " AND NumeOper =" & NOPER%
   FEC% = CVINT(VALOBADA("SECOPER", "FECH_REVIS", CONDI$, "NOMESS"))
   '
   If FEC% < 1 Then
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
   End If
   '
   NPRO% = CVI(Mid$(RESECO$(List3.ListIndex), 149, 2))
   PREC# = CVD(Mid$(RESECO$(List3.ListIndex), 151, 8))
   MONEI% = CVI(Mid$(RESECO$(List3.ListIndex), 159, 2))
   If MONEI% < 1 Then MONEI% = 1
   FECOI = CVI(Mid$(RESECO$(List3.ListIndex), 161, 2))
   PCLX = CVS(Mid$(RESECO$(List3.ListIndex), 99, 4))
   '
   Text37.TEXT = CSTRING$(MKD$(PREC#), 3, 11, 4, 2)
   Text34.TEXT = CSTRING$(MKS$(PCLX), 3, 5, 1, 2)
   '
   If MONEI% > 0 Then
     If MONEI% < 256 Then
       MONEX$ = AJUSTI$(ECOARCH$("TAMONED", Chr$(MONEI%)), 48) + Str$(MONEI%)
       Label46.Caption = MONEX$
     End If
   End If
   '
'   Label63 = FECHATEX$(FECOI) corresponde al archivo de memoria tecnica
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
   If MODIFIC%(0) > 0 Then GoSub 90: Call GRAGENITEM
   If MODIFIC%(1) > 0 Then GoSub 90: Call GRAESPECI
   If MODIFIC%(2) > 0 Or (SSTab1.Tab = 1 And GRABATODO% = 1) Then Command9.SetFocus: Call GRAFORMULA
   If MODIFIC%(3) > 0 Then GoSub 90: Call GRAHORUTA
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
      Call REPLA(YY$, MKD$(XVALO(Label90) + XVALO(Label82)), 9, 8)
      Call REPLA(YY$, MKD$(XVALO(Label88) + XVALO(Label83)), 17, 8)
      Call REPLA(YY$, MKD$(XVALO(Label86) + XVALO(Label84)), 25, 8)
      Call REPLA(YY$, MKD$(XVALO(Label85)), 33, 8)
      Call REPLA(YY$, MKD$(XVALO(Label75) + XVALO(Label72)), 41, 8)
      Call REPLA(YY$, MKD$(XVALO(Label79) + XVALO(Label74)), 49, 8)
      Call REPLA(YY$, MKD$(0), 57, 8)
      Call REPLA(YY$, MKD$(XVALO(Label80)), 65, 8)
      Call REPLA(YY$, MKD$(XVALO(Label70)), 73, 8)
      Call REPLA(YY$, MKD$(XVALO(Label65)), 81, 8)
      Call REPLA(YY$, MKD$(XVALO(Label96)), 89, 8)
      Call REPLA(YY$, MKD$(XVALO(Label103)), 97, 8)
      Call REPLA(YY$, MKD$(XVALO(Label112)), 105, 8)
      Call REPLA(YY$, MKD$(XVALO(Label116)), 113, 8)
      Call REPLA(YY$, MKS$(XVALO(Label117)), 121, 4)
      Call REPLA(YY$, MKS$(XVALO(Label119)), 125, 4)
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
   'SOLICITUD DE DOSIVAC QUE PUEDAN OTORGAR SOLO DERECHO DE MODIFICAR SEGUNDA DESCRIPCION
   If DERACCE%("MODITMAS/NOMESS", "Modificación Datos Maestro de Stock") < 2 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      'SI ESTA CONFIGURADA LA DESCRIPCION DOBLE
      If CONTROL("MASTER", "DESLARGA") <> "SI" And InStr(EPAC$, "DOSIVA") > 0 Then
         'VALIDA EL DERECHO
         If DERACCE%("MODI2DES", "Modificación Segunda Descripción") > 1 Then
             Call ACTUREGISTRO("MASTER", "Descridos", "codint = " & CI%, Left$(Text38, 64), REGAF&, "NOMESS")
             Exit Sub
         End If
      End If
   End If
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") < 2 Then
      Exit Sub
   End If
   '
   If INCLUALIAS% = 1 Then
     ALIA$ = TRIM$(Text45)
     CDIXI% = CODINT%(ALIA$)
     If CDIXI% > 0 Then
       If CDIXI% <> CI% Then
         Call MENSERR(24, "Imposible Registrar.\Alias No Válido o Pre-Existente.")
         Exit Sub
       End If
     End If
   End If
   '
   If XVALO(Text50) > 12 Then
      Call MENSERR(24, "La Cantidad de Decimales no Puede Ser Mayor a 12")
      Exit Sub
   End If
   '
   If YACAR%(0) < 1 Then    ' SOLO CAMBIO LA DESCRIPCION O EL NRO DE PLANO
      Screen.MousePointer = 11
      MODESCRI% = 0
      CIXI% = CODINT%(Text1.TEXT) ' SE CAMBIA PARA VOLVER A LEER EL CODIGO INTERNO
      DE$ = DESCRIT0$(CIXI%): DE2$ = DESCRIT2$(CIXI%)
      If TRIM$(Text3) <> "" Then
        If TRIM$(Text3) <> DE$ Or TRIM$(Text38) <> DE2$ Or TRIM$(Text45) <> NUPLANO$(CIXI%) Then
          DE$ = TRIM$(Text3): DE2$ = TRIM$(Text38)
          MODESCRI% = 1
          '
          CONDI$ = "codint = " & CIXI%
          Call ACTUREGISTRO("MASTER", "Descripcion", CONDI$, Left$(DE$, 64), REGAF&, "NOMESS")
          Call ACTUREGISTRO("MASTER", "Descridos", CONDI$, Left$(DE2$, 64), REGAF&, "NOMESS")
          Call ACTUREGISTRO("MASTER", "Nplano", CONDI$, Left$(Text45, 24), REGAF&, "NOMESS")
          '
'            STRSQL = "Update Master set Descripcion= '" & Left$(DE$, 64)
'            STRSQL = STRSQL & "', Descridos= '" & Left$(DE2$, 64)
'            STRSQL = STRSQL & "', Nplano= '" & Left$(Text45, 24)
'            STRSQL = STRSQL & "' where codint= " & CI%
'             FLEXCONN.Execute (FILTSQL$(STRSQL))
          
         ' Master.FindFirst ("CODINT = " & CI%)
         ' Master.Edit
         ' Master!DESCRIPCION = Left$(DE$, 64)
         ' Master!DESCRIDOS = Left$(DE2$, 64)
         ' Master!NPLANO = Left$(Text45, 24)
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
   'VALIDACION DE ATRIBUTO A SI REPITE
   If CONTROL("AMASTO", "REPATRIA") = "SI" And TRIM$(Text23.TEXT) <> "" Then
        TTALLE$ = TRIM$(Text23.TEXT)
        CONDI$ = "TALLE = '" & TTALLE$ & "' AND CODINT <> " & CI%
        If CantRegistros("MASTER", CONDI$) > 0 Then
           Call MENSERR(24, "Atributo A :" & TTALLE$ & " Ya Existente en otro Código")
           Text23 = VALOBADA("MASTER", "TALLE", "CODINT = " & CI%, "NOMESS")
        End If
   End If
   '
   'VALIDACION DE ATRIBUTO A SI REPITE
   If CONTROL("AMASTO", "REPATRIB") = "SI" And TRIM$(Text22.TEXT) <> "" Then
        TCOLOR$ = TRIM$(UCase$(Text22.TEXT))
        CONDI$ = "COLOR = '" & TCOLOR$ & "' AND CODINT <> " & CI%
        If CantRegistros("MASTER", CONDI$) > 0 Then
           Call MENSERR(24, "Atributo B :" & TCOLOR$ & " Ya Existente en otro Código")
           Text22 = VALOBADA("MASTER", "COLOR", "CODINT = " & CI%, "NOMESS")
           Exit Sub
        End If
   End If

   Screen.MousePointer = 11
   '
   MODESCRI% = 0
   CIXI% = CODINT%(Text1.TEXT) ' SE CAMBIA PARA VOLVER A LEER EL CODIGO INTERNO
   
   Dim RMaster As ADODB.Recordset
   Set RMaster = New ADODB.Recordset
   RMaster.CursorType = adOpenDynamic
   RMaster.LockType = adLockPessimistic
   strsql = ("Select * from Master where codint=") & CIXI%
   RMaster.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
   If Not (RMaster.EOF And RMaster.BOF) Then
      'Master.FindFirst ("CODINT = " & CI%)
      'Master.Edit
   '
   DE$ = DESCRIT0$(CIXI%): DE2$ = TRIM$(Text38)
   If TRIM$(Text3) <> "" Then
     If TRIM$(Text3) <> DE$ Then
       DE$ = TRIM$(Text3)
       MODESCRI% = 1
       RMaster!Descripcion = Left$(DE$, 64)
     End If
   End If
   '
   RMaster!DESCRIDOS = Left$(DE2$, 64)
   TTITPE$ = TRIM$(Text24.TEXT)
   PESU = XVALO(TTITPE$)
   If InStr(TTITPE$, "/") > 0 Then
      TTQP$ = TTITPE$
      TTQP$ = REPLACAR$(TTQP$, "/", ".")
      PESU = (-1) * Abs(XVALO(TTQP$))
   End If
   RMaster!PESUNI = PESU
   '
   TIMA% = XVALO(Text6.TEXT)
   RMaster!TIMATE = TIMA%
   '
   FAMI$ = Left$(Text48.TEXT, 2)
   RMaster!Familia = FAMI$
   '
   GRUIVA% = XVALO(Text35.TEXT)
   RMaster!GRUPIVA = GRUIVA
   '
   CRIT% = XVALO(Text8.TEXT)
   RMaster!CRITREP = CRIT%
   '
   MONEI% = XVALO(Text13.TEXT)
   RMaster!Monecos = MONEI%
   '
   TCOLOR$ = TRIM$(Text22.TEXT)
   RMaster!COLOR = Left$(TCOLOR$, 32)
   XCI% = CCOLOR%(TCOLOR$)
   '
   TTALLE$ = TRIM$(Text23.TEXT)
   RMaster!TALLE = Left$(TTALLE, 32)
   XCI% = CTALLE%(TTALLE$)
   '
   CO_INTRERNA$ = TRIM$(Text44.TEXT)
   RMaster!Cod_Interna = Left$(CO_INTRERNA$, 64)
   '
   RMaster!Revision = Left$(TRIM$(Text36), 16)
   FREVI% = FECHANUM(Text25.TEXT)
   RMaster!FEREVIS = CVDAT(FREVI%)
   '
   STOMINI = XVALO(Text18)
   RMaster!STOMIN = STOMINI
   '
   'GUARDO EN UNA VARIABLE EL VALOR ANTERIOR
   LOTEECOANT = LOTESTAN(CI%)
   '
   LOREPO = XVALO(Text19)
   RMaster!LOREPOS = LOREPO
   '
   STOMAX = XVALO(Text20)
   RMaster!STOMAX = STOMAX
   '
   UNID$ = TRIM$(Text4.TEXT)
   RMaster!UMEDIDA = Left$(UNID$, 4)
   '
   If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 0 Then
     COU = XVALO(Text12.TEXT)
     RMaster!COSUNI = COU
     '
     FECOSTX = Text14.TEXT
     If FECOSTX = "" Then FECOSTX = Now
     On Error Resume Next
     RMaster!FECOSTO = CDate(FECOSTX)
     On Error GoTo 0
     '
   End If
   
   PORNA = XVALO(Text28.TEXT)
   RMaster!PORNACIO = PORNA
   '
   NPRO% = XVALO(Text10.TEXT)
   RMaster!PROVHABI = NPRO%
   NPROA% = CVI(Mid$(REMAS$, 123, 2))
   '
   ' ESTO ATAJA LA POSIBILIDAD DE QUE EL VECTOR
   ' NO TENGA AL MENOS UN ELEMENTO
   On Error Resume Next
   LGPHABIT% = UBound(PHABIT)
   On Error GoTo 0
   '
   If CIXI% > LGPHABIT% Then
      ReDim Preserve PHABIT(CIXI%)
   End If
   '
   PHABIT%(CIXI%) = (-1)
   If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
     PHABIT%(CIXI%) = NPRO%
   End If
   '
   GRUCO% = XVALO(Text16.TEXT)
   RMaster!Grucoven = GRUCO%
   '
   If Abs(XVALO(Text28)) > 320 Then
       Call MENSERR(24, "% Gastos de Nacionalización No Válido.")
     Else
       FCPI% = Int(100 * XVALO(Text28.TEXT) + 0.5)
'       Call REPLA(REMAS$, MKI$(FCPI%), 127, 2)
   End If
   '
   '\\\OT-06-0105-WAR-23/03/06///
   DECOAU1% = 0
   If Check3.Value = 1 Then DECOAU1% = 1
   On Error Resume Next
   RMaster!DesCompAut = DECOAU1%
   On Error GoTo 0
   '\\\OT-06-0105-WAR-FIN///
   
   '\\\OT-08-0135-OMAR-04/04/08///
   TRAZ% = XVALO(Label58)
   On Error Resume Next
   RMaster!TRAZABLE = TRAZ%
   On Error GoTo 0
   '\\\OT-08-0135-OMAR-FIN///
   '
   '\\\OT-5-0633-FG-21/09/05///
   PESME = XVALO(Text46.TEXT)
   On Error Resume Next
   RMaster!PESMETRO = PESME
   On Error GoTo 0
   '\\\OT-5-0633-FG-FIN///
   '
   '\\\OT-05-0720-WAR-25/11/05///
   LOINDI = XVALO(Text40.TEXT)
   On Error Resume Next
   RMaster!LINDIVIS = LOINDI
   On Error GoTo 0
   '\\\OT-05-0720-WAR-FIN///
   '
   LOINDIVTA = XVALO(Text49.TEXT)
   On Error Resume Next
   RMaster!LINDIVEN = LOINDIVTA
   On Error GoTo 0
   '
   LTC = XVALO(Text51.TEXT)
   On Error Resume Next
   RMaster!LeTiCom = LTC
   On Error GoTo 0
   '
   'PERIODO DE STOCKS  MINIMO
   PSM$ = TRIM$(Text52.TEXT)
   On Error Resume Next
   RMaster!Autono_Proy = PSM$
   On Error GoTo 0
   '
30 MODIPLAN% = 0: PLANUE$ = ""
   On Error Resume Next
   PLANUE$ = TRIM$(RMaster!NPLANO)
   If PLANUE$ <> TRIM$(Left$(Text45, 24)) Then
     MODIPLAN% = 1
     UPDAPLANOS% = 1
   End If
   
   RMaster!NPLANO = TRIM$(Left$(Text45, 24))
'   RMaster!Especificacion = Text43 ' NO CORRESPONDE QUE GUARDE EN ESTA RUTINA
   '
   RMaster!CANDECI = TRIM$(Text50.TEXT)
   '
   'Estado de movimientos (Gabal- Dato segun proceso Rutina: EstadoMovim)
   On Error Resume Next
   RMaster!Stat_Movi = XVALO(Label71) - 1
   On Error GoTo 0
   '
   On Error Resume Next
   RMaster.Update
   RMaster.Close
   Set RMaster = Nothing
   On Error GoTo 0
   UPDATMASTER% = 1
   End If
   'Fin actualizacion
   '
   If NPRO% <> NPROA% Then  ' cambio de proveedor habitual
     ZZX$ = REGBLAN$("RECAMPH")
     Call REPLA(ZZX$, MKI$(CIXI%), 1, 2)
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
   If MODESCRI% > 0 Or MODIPLAN% > 0 Then
      Call GENECOMAS
   End If
   '
   
   If CONTROL("MASTER", "FORPATRO") = "SI" And LOTEECOANT <> LOREPO Then
      Call ACTUFORMUPAT(CI%, LOTEECOANT, LOREPO)
   End If
   
   
   
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
   If Len(Label63) > 256 Then
      COMUNI ("Longitud de Memoria Tecnica Supera el Maximo Permitido")
      Exit Sub
   End If
   '

   Screen.MousePointer = 11
   
   'Master.FindFirst ("CODINT = " & CI%)
   'Master.Edit
   'Master!ESPECIFICACION = Text43
   'Master.Update
   '
   '
'   STRSQL = "Update Master set Especificacion = '" & Text43 & "' where codint = " & CI%
'   FLEXCONN.Execute STRSQL, cdadreg
   '
   'SE MODIFICA LA GRABACION CON ACTUREGISTRO POR QUE TIENE CAPTURADO EL ERROR QUE ESTA
   'PRODUCIENDO EN LA LINEA DE ARRIBA CUANDO SE INTENTA GRABAR UN TEXTO CON APOSTROFE
   Call ACTUREGISTRO("Master", "Especificacion", "codint = " & CI%, Text43, REG&)
   '
   Call ACTUREGISTRO("Master", "ArchiPdf", "codint = " & CI%, Label63, REG&)
   Call ACTUREGISTRO("Master", "Clas_ATC", "codint = " & CI%, Text81, REG&)
   Call ACTUREGISTRO("Master", "Vers_BC ", "codint = " & CI%, Text82, REG&)
   Call ACTUREGISTRO("Master", "FeVig_BC", "codint = " & CI%, FETEXCOR1$(FECHANUM(Text83)), REG&)
   Call ACTUREGISTRO("Master", "FeRev_BC", "codint = " & CI%, FETEXCOR1$(FECHANUM(Text84)), REG&)
   '
   'GRAABAR P.O.E.`S
   '****************
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCUREL", "MARBORRA", "COD_INTE = " & CI%, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For I& = 1 To List2.ListCount
      Z$ = List2.List(I& - 1)
      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
      rutax$ = TRIM$(Mid$(Z, 32, 512))
10    RNDX$ = RANDSTRING$(8)
      If CantRegistros("DOCUREL", "DESCRIPCION = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10

      Call CreaRegistro("DOCUREL", "DESCRIPCION", RNDX$, "NOMESS")
      Call ACTUREGISTRO("DOCUREL", "COD_INTE", "DESCRIPCION = '" & RNDX$ & "'", CI%, REG&)
      Call ACTUREGISTRO("DOCUREL", "MARBORRA", "DESCRIPCION = '" & RNDX$ & "'", 0, REG&)
      Call ACTUREGISTRO("DOCUREL", "RUTA", "DESCRIPCION = '" & RNDX$ & "'", rutax$, REG&)
      Call ACTUREGISTRO("DOCUREL", "DESCRIPCION", "DESCRIPCION = '" & RNDX$ & "'", DESCRI$, REG&)
   Next I&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCUREL", "MARBORRA = 1 AND COD_INTE = " & CI%, REGAF&, "NOMESS")
   
   Screen.MousePointer = 1
    
   
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
      REVIFORMU = HOY(HOS$)
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
   
   Call ABRESECOP 'llama a abreconexion
   'sql
   Dim RSECOPER As ADODB.Recordset
   Set RSECOPER = New ADODB.Recordset
   'RSECOPER.CursorType = adOpenDynamic
   RSECOPER.CursorType = adOpenKeyset
   RSECOPER.LockType = adLockPessimistic
   strsql = ("Select * from Secoper where CodIConj=") & CI%
   RSECOPER.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
   With RSECOPER
      strsql_U = ("Update Secoper set StatOper = -1 where StatOper>0 and CodIConj=") & CI%
      FLEXCONN.Execute strsql_U, cdadreg
   
   
   '
   'With Secoper
   '  YAPRIMO% = 0
   '  On Error Resume Next
   '  .MoveFirst
   '  If Err > 0 Then GoTo 65
   '  On Error GoTo 0
     '
   '  ABUS$ = "CodIConj =" & CI% & " AND StatOper > 0"
   '  On Error Resume Next
62 '  If YAPRIMO% = 0 Then
   '      .FindFirst ABUS$
   '      YAPRIMO% = 1
   '    Else
   '      .FindNext ABUS$
   '  End If
   '  If .NoMatch = False Then
   '    .Edit
   '    !statoper = (-1)
   '    .Update
   '    GoTo 62
   '  End If
   'End With
   '
65 'On Error GoTo 0
   
      
    For LL% = 1 To List3.ListCount
    NOPYI% = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
    CANPRE = 0: CATPRE% = 0
     ' VER ERNESTO
     '.FindFirst "CODICONJ =" & CI% & " and NumeOper = " & NOPYI%
     'If .NoMatch = False Then
     Dim MIRS As ADODB.Recordset
     strsql_S = ("Select * from Secoper where CODICONJ=") & CI% & " and NumeOper=" & NOPYI%
     Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql_S))
     If Not (MIRS.EOF And MIRS.BOF) Then
       CANPRE = MIRS!NUOPER_PREP
       CATPRE% = MIRS!CATOPER_PREP
     End If
     'End If
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !CODICONJ = CI%
     !CODXCONJ = CODEXT$(CI%)
     !NumeOper = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
     !AltOper = Asc(Mid$(RESECO$(LL% - 1), 88, 1))
     !DescOper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
     !StatOper = 1
     !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
     !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
     !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
     !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
     !CodMaq4 = TRIM$(Mid$(RESECO$(LL% - 1), 182, 6))
     !CodMaq5 = TRIM$(Mid$(RESECO$(LL% - 1), 188, 6))
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
     !CantiOps2 = CVS(Mid$(RESECO$(LL% - 1), 178, 4))
     !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
     !CategOps2 = Asc(Mid$(RESECO$(LL% - 1), 177, 1))
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
        FE% = HOY(HO$)
     !FECH_REVIS = CVDAT(FE%)
     !USUA_REVIS = USERNAME$
     
     .Update
   
   Next LL%
   RSECOPER.Close
   Set RSECOPER = Nothing
   End With
      strsql_D = ("Delete Secoper where StatOper=-1 and CodIConj=") & CI%
      FLEXCONN.Execute strsql_D, cdadreg
   
   
   '
63 '.FindFirst "CODICONJ =" & CI%
   'If .NoMatch = False Then
   '  If !statoper = (-1) Then
   '    .Delete
   '    GoTo 63
   '  End If
   'End If
   '
64 '.FindNext "CODICONJ =" & CI%
   'If .NoMatch = False Then
   '  If !statoper = (-1) Then
   '    .Delete
   '    GoTo 64
   '  End If
   'End If
   '
   '.Close
   'End With   ' secoper
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

Private Sub Listad_Click()
   Call OPLIMAS07.Command2_Click
End Sub

Private Sub MaePro_Click()
   Call CONECRUN("ARCHIG07/AMAPRO", "Padrón Maestro de Proveedores")
End Sub

Private Sub NumPro_Click()
   Call OPLIMAS07.Command5_Click
End Sub

Private Sub Pad_Click()
   Call Command45_Click
End Sub

Private Sub Otro_Click()
   Call ACCDIR("AMASTO")
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
Call OPLIMAS07.Command7_Click
End Sub

Private Sub SustCod_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   '
   For KKII% = 0 To 15: YACAR%(KKII%) = 0: Next KKII%
   '
   
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   FORSELEC.ABUSQUE = Text1
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
   Label33 = RASOCLI$((-1) * XVALO(Text10))
   MODIFIC%(0) = 1
End Sub

Private Sub Text10_DblClick()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(XVALO(Text10.TEXT))))
   End If
End Sub


Private Sub Text10_LostFocus()
   PROVEHABI = XVALO(Text10.TEXT)
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
     MONEI% = XVALO(Text13.TEXT)
     COSUDOL = XVALO(Text12)
     If MONEI% > 1 Then
         VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100)
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
   Text12.TEXT = CSTRING$(MKS$(XVALO(Text12.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text13_Change()
   VALONU = XVALO(Text13.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text13.TEXT = "": Exit Sub
   End If
   Label35.Caption = ECOARCH$("TAMONED", Chr$(VALONU))
   '
   MONEI% = XVALO(Text13.TEXT)
   'If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = XVALO(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = XVALO(Text21) / (TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100))
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
      MONEI% = XVALO(Text13)
      COSUDOL = XVALO(Text12)
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
   If XVALO(Text15(Index).TEXT) > 999 Then Text15(Index).TEXT = "999"
   If XVALO(Text15(Index).TEXT) < 0 Then Text15(Index).TEXT = "  0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = XVALO(Text15(Index).TEXT) + XVALO(Text17(Index).TEXT) / 60 + XVALO(Text26(Index).TEXT) / 3600
     If Index = 2 Then
        TPAR = 60 * TPAR / XVALO(Text7.TEXT)
        Call CALCADEN
     End If
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   'Call CALCOSOPE
End Sub

Sub CALCADEN()
   CADEN = 0
   TSTANBA = XVALO(Text15(2)) + XVALO(Text17(2)) / 60 + XVALO(Text26(2)) / 3600
   If XVALO(Text7) > 0.5 Then TSTANBA = TSTANBA / XVALO(Text7)
   If TSTANBA > 0 Then CADEN = 1 / TSTANBA
   If CADEN >= 1000 Then CADEN = Int(CADEN + 0.5)
   Label50 = FORMATNUM(CADEN, "F7.1")
End Sub
'
Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index).TEXT)
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(XVALO(Text15(Index).TEXT)), 3, 3, 0, 2))
   If TTTT$ = "" Then TTTT$ = "0"
   While Len(TTTT$) < 3: TTTT$ = " " + TTTT$: Wend
   Text15(Index).TEXT = TTTT$
End Sub

Private Sub Text17_Change(Index As Integer)
   If XVALO(Text17(Index).TEXT) > 59 Then Text17(Index).TEXT = "59"
   If XVALO(Text17(Index).TEXT) < 0 Then Text17(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = XVALO(Text15(Index).TEXT) + XVALO(Text17(Index).TEXT) / 60 + XVALO(Text26(Index).TEXT) / 3600
     If Index = 2 Then
        TPAR = 60 * TPAR / XVALO(Text7.TEXT)
        Call CALCADEN
     End If
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   'Call CALCOSOPE
End Sub

Private Sub Text17_GotFocus(Index As Integer)
   Text17(Index).SelStart = 0
   Text17(Index).SelLength = Len(Text17(Index).TEXT)
End Sub

Private Sub Text17_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(XVALO(Text17(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text17(Index).TEXT = TTTT$
End Sub

Private Sub Text18_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text18_LostFocus()
   Text18.TEXT = CSTRING$(MKS$(XVALO(Text18.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
   If XVALO(Text20.TEXT) < XVALO(Text19.TEXT) Then
     Text20.TEXT = CSTRING$(MKS$(XVALO(Text19.TEXT)), 3, 10, 1, 2)
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19.TEXT = CSTRING$(MKS$(XVALO(Text19.TEXT)), 3, 10, 1, 2)
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
   Text20.TEXT = CSTRING$(MKS$(XVALO(Text20.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text21_Change()
   If FOCOREP% = 2 Then
     MONEI% = XVALO(Text13.TEXT)
     On Error Resume Next
     COSUDOL = XVALO(Text21) / (TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100))
     On Error GoTo 0
     Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text21_GotFocus()
   FOCOREP% = 2
End Sub

Private Sub Text21_LostFocus()
   Text21.TEXT = CSTRING$(MKS$(XVALO(Text21.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text22_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text22_DblClick()
   DETABLA$ = UCase$(Label10(3))
   PPXX% = InStr(DETABLA$, "- ")
   If PPXX% > 0 Then DETABLA$ = Left$(DETABLA$, PPXX% - 1) + Mid$(DETABLA$, PPXX% + 2)
   Call SELECHO("TACOLOR/" + DETABLA$)
   COLONUE$ = VALACT1$("TACOLOR")
   If COLONUE$ <> "" Then
      Text22.TEXT = COLONUE$
   End If
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_DblClick()
   DETABLA$ = UCase$(Label10(2))
   PPXX% = InStr(DETABLA$, "- ")
   If PPXX% > 0 Then DETABLA$ = Left$(DETABLA$, PPXX% - 1) + Mid$(DETABLA$, PPXX% + 2)
   Call SELECHO("TATALLES/" + DETABLA$)
   TALLENUE$ = VALACT1$("TATALLES")
   If TALLENUE$ <> "" Then
      Text23.TEXT = TALLENUE$
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_LostFocus()
   CANDE% = 5: If XVALO(Text24.TEXT) < 10 Then CANDE% = 7
   Text24.TEXT = CSTRING$(MKS$(XVALO(Text24.TEXT)), 3, 10, CANDE%, 2)
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_DblClick()
   Call Command20_Click
End Sub

Private Sub Text26_Change(Index As Integer)
   If XVALO(Text26(Index).TEXT) > 59.9999 Then Text26(Index).TEXT = "59.9999"
   If XVALO(Text26(Index).TEXT) < 0 Then Text26(Index).TEXT = "00.0000"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = XVALO(Text15(Index).TEXT) + XVALO(Text17(Index).TEXT) / 60 + XVALO(Text26(Index).TEXT) / 3600
     If Index = 2 Then
       TPAR = 60 * TPAR / XVALO(Text7.TEXT)
       Call CALCADEN
     End If
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   'Call CALCOSOPE
End Sub

Private Sub Text26_GotFocus(Index As Integer)
   Text26(Index).SelStart = 0
   Text26(Index).SelLength = Len(Text26(Index).TEXT)
End Sub

Private Sub Text26_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(XVALO(Text26(Index).TEXT)), 3, 7, 4, 2))
   While Len(TTTT$) < 7: TTTT$ = "0" + TTTT$: Wend
   If TTTT$ = "0000000" Then TTTT$ = "00.0000"
   Text26(Index).TEXT = TTTT$
End Sub

Private Sub Text28_Change()
   MONEI% = XVALO(Text13.TEXT)
   'If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = XVALO(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = XVALO(Text21) / (TICAMONE(MONEI%) * (1 + XVALO(Text28) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_LostFocus()
  'If XVALO(Text13) <= 1 Then Text28 = ""
  Text28 = AJUSTD$(FORMATNUM$(XVALO(Text28), "F7.2"), 8)
End Sub

Private Sub Text29_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 75, 4) = MKS(XVALO(Text29.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text29_GotFocus()
   Text29.SelStart = 0
   Text29.SelLength = Len(Text29.TEXT)
End Sub

Private Sub Text29_LostFocus()
   Text29.TEXT = CSTRING$(MKS$(XVALO(Text29.TEXT)), 3, 4, 1, 2)
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text30_Change(Index As Integer)
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     If Index = 0 Then
         Mid$(VDEF$, 83, 4) = MKS(XVALO(Text30(Index).TEXT))
       ElseIf Index = 1 Then
         Mid$(VDEF$, 178, 4) = MKS(XVALO(Text30(Index).TEXT))
     End If
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text30_GotFocus(Index As Integer)
   Text30(Index).SelStart = 0
   Text30(Index).SelLength = Len(Text30(Index).TEXT)
End Sub

Private Sub Text30_LostFocus(Index As Integer)
   Text30(Index).TEXT = CSTRING$(MKS$(XVALO(Text30(Index).TEXT)), 3, 5, 2, 2)
End Sub

Private Sub Text31_Change()
   Command26.Enabled = False
   Command27.Enabled = False
   Label40.Caption = ""
   Label41.Caption = ""
   Label42.Caption = ""
   Text32 = ""
   CICOMPO% = CODINT%(Text31)
   If CICOMPO% > 0 Then
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     Label42.Caption = FORMATNUM$(PESMETRO(CICOMPO%), "F7.3")
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Command26.Enabled = True
     Command27.Enabled = True
   End If
End Sub

Private Sub Text33_GotFocus(Index As Integer)
   Text33(Index).SelStart = 0
   Text33(Index).SelLength = Len(Text33(Index).TEXT)
End Sub
'
Private Sub Text34_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 99, 4) = MKS(XVALO(Text34.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   'Call CALCOSOPE
End Sub

Private Sub Text34_GotFocus()
   Text34.SelStart = 0
   Text34.SelLength = Len(Text34.TEXT)
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
   If PORLARGO% = 1 Then
     On Error Resume Next
     Text41.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   Text32.SelStart = 0
   Text32.SelLength = Len(Text32.TEXT)
End Sub

Private Sub Text33_Change(Index As Integer)
   VALONU = XVALO(Text33(Index).TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text33(Index).TEXT = "": Exit Sub
   End If
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     If Index = 0 Then
          Mid$(VDEF$, 87, 1) = Chr$(VALONU)
        ElseIf Index = 1 Then
          Mid$(VDEF$, 177, 1) = Chr$(VALONU)
     End If
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Label51(Index).Caption = ECOARCH$("CATOPER", Chr$(VALONU))
End Sub

Private Sub Text33_DblClick(Index As Integer)
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33(Index).TEXT = VALONUE$
   End If
End Sub

Private Sub Text35_Change()
   VALONU = XVALO(Text35.TEXT)
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
   MODIFIC%(0) = 1
End Sub

Private Sub Text37_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 151, 8) = MKD(XVALO(Text37))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = XVALO(Mid$(Label46, 49, 6))
   TERPESOS = XVALO(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_LostFocus()
   Text37.TEXT = CSTRING$(MKS$(XVALO(Text37.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text34_LostFocus()
   Text34.TEXT = CSTRING$(MKS$(XVALO(Text34.TEXT)), 3, 5, 1, 2)
End Sub

Private Sub Text38_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text39_GotFocus()
   Text39.SelStart = 0
   Text39.SelLength = Len(Text39.TEXT)
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
Private Sub Text49_Change()
   MODIFIC%(0) = 1
End Sub


Private Sub Text44_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text46_Change()
   MODIFIC%(0) = 1
End Sub

'Private Sub TEXT41_CHANGE()
'   Label93 = TRIM$(CSTRING$(MKD$(XVALO(Label85) * XVALO(Text41) / 100), 3, 12, DECICALCOS%, 2))
'End Sub
Private Sub Text41_GotFocus()
   Text41.SelStart = 0
   Text41.SelLength = Len(Text41.TEXT)
End Sub

Private Sub Text43_Change()
   MODIFIC%(1) = 1
End Sub

Private Sub Text45_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text47_GotFocus()
   Text47.SelStart = 0
   Text47.SelLength = Len(Text47)
End Sub

Private Sub Text47_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Text47_LostFocus
   End If
End Sub

Private Sub Text47_LostFocus()
   Text47 = FORMATNUM(XVALO(Text47), "F7.1")
   CADENSTA = XVALO(Text47)
   If CADENSTA >= 0.1 Then
      LSTD = XVALO(Text7): If LSTD < 1 Then LSTD = 1
      TISTA = LSTD * 1 / CADENSTA
      TSTT$ = HORMINSEDIEM$(3600 * TISTA)
      Text15(2).TEXT = Left$(TSTT$, 3)
      Text17(2).TEXT = Mid$(TSTT$, 5, 2)
      Text26(2).TEXT = Mid$(TSTT$, 8, 7)
    End If
    Text47.Visible = False
End Sub

Private Sub Text48_Change()
   TAFAM$ = TRIM$(Text48.TEXT)
   MODIFIC%(0) = 1
   If TAFAM$ = "" Then
      Text48.TEXT = ""
      Label64 = ""
      Exit Sub
   End If
   Label64.Caption = ECOARCH$("TAFAMIL", TAFAM$)
   MODIFIC%(0) = 1

End Sub

Private Sub Text48_DblClick()
   Call Command55_Click
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

Private Sub Text5_lostfocus()
   ILISTA& = List3.ListIndex
   For UI& = 1 To List3.ListCount
     If XVALO(Left$(List3.List(UI& - 1), 4)) = XVALO(Label55) Then
       If XVALO(Mid$(List3.List(UI& - 1), 6, 2)) = XVALO(Label52) Then
         TTXT$ = List3.List(UI& - 1)
         List3.RemoveItem (UI& - 1)
         Call REPLA(TTXT$, Text5.TEXT, 10, 48)
         List3.AddItem TTXT$, UI& - 1
       End If
     End If
   Next UI&
   On Error Resume Next
   List3.ListIndex = ILISTA&
   On Error GoTo 0
End Sub

Private Sub Text50_Change()
  MODIFIC%(0) = 1
End Sub

Private Sub Text51_Change()
  MODIFIC%(0) = 1
End Sub

Private Sub Text52_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_Change()
   VALONU = XVALO(Text6.TEXT)
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
   VALONU = XVALO(Text16.TEXT)
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
     CABAS = XVALO(Text7.TEXT)
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
   MODIFIC%(3) = 1
   '
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub Text8_Change()
   VALONU = XVALO(Text8.TEXT)
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

Private Sub Text81_Change()
    MODIFIC%(1) = 1
End Sub



Private Sub Text81_GotFocus()

     If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub

End Sub


Private Sub Text82_Change()
    MODIFIC%(1) = 1
End Sub

Private Sub Text82_GotFocus()
    If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub
End Sub

Private Sub Text83_Change()
    MODIFIC%(1) = 1
End Sub

Private Sub Text83_GotFocus()
    If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub
End Sub

Private Sub Text84_Change()
   MODIFIC%(1) = 1
End Sub


Private Sub Text84_GotFocus()
     '
    If DERACCE%("MODFOPAT/NOMESS", "Modificar Formula Patrón") < 2 Then Call Text1.SetFocus: Exit Sub     '

End Sub

Private Sub Text9_Change(Index As Integer)
   If List3.ListIndex >= 0 Then
     If TRIM$(Text9(Index).TEXT) = "" Or ECOARCH$("PADMAQ", Text9(Index).TEXT) <> "" Then
        If Index = 0 Then
            Label43.Caption = ECOARCH$("PADMAQ", Text9(Index).TEXT)
            Mid$(RESECO$(List3.ListIndex), 53, 6) = AJUSTI$(Text9(Index).TEXT, 6)
          ElseIf Index < 4 Then
            Mid$(RESECO$(List3.ListIndex), 97 + 6 * Index, 6) = AJUSTI$(Text9(Index).TEXT, 6)
          Else
            Mid$(RESECO$(List3.ListIndex), 158 + 6 * Index, 6) = AJUSTI$(Text9(Index).TEXT, 6)
        End If
     End If
   End If
   MODIFIC(3) = 1
End Sub

Private Sub Text9_DblClick(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
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
     ALTY% = Asc(Mid$(RESECO$(KU% - 1), 88, 1))
     NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
     DOPE$ = Mid$(RESECO$(KU% - 1), 5, 48)
     SECO$ = NOPEX$ + "     " + DOPE$
     If ALTY% > 0 Then
       Mid$(SECO$, 5, 5) = "-" + TRIM$(Str$(ALTY%))
     End If
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
         x$ = REGLEIDO$("PLACONTR", RESE&)
         Call REPLA(x$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = x$
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
         x$ = REGLEIDO$("PUEPUNTO", RESE&)
         Call REPLA(x$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = x$
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
         x$ = REGLEIDO$("ESPECIFI", RESE&)
         Call REPLA(x$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = x$
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
              x$ = REGLEIDO$("MASTER", I& + 1)
              CODES$ = Mid$(x$, 77, 16)
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
                 Call REPLA(x$, CODIS$, 77, 16)
                 Call GRAREG("MASTER", x$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              x$ = REGLEIDO$("INVERT", I& + 1)
              CODES$ = Mid$(x$, 1, 16)
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
                 Call REPLA(x$, CODIS$, 1, 16)
                 Call GRAREG("INVERT", x$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("INVERT")
           '
        End If
        '
        NM% = ULTREG&("MASTER") - 1
        If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
              x$ = REGLEIDO$("MASTER", 1&)
              Call REPLA(x$, MKI$(NM%), 1, 2)
              Call GRAREG("MASTER", x$, 1&)
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
            x$ = REGLEIDO$("INVERT", II& + 1)
            COD$ = Left$(x$, 16): CII% = CVI(Mid$(x$, 17, 2))
            If Mid$(COD$, 16, 1) = Chr$(96) Then
               If CII% > 0 Then
                  If CII% <= NUMAS% Then
                     REGMA& = CII% + 1
                     y$ = REGLEIDO$("MASTER", REGMA&)
                     If Mid$(y$, 77, 15) = Left$(COD$, 15) Then
                        Call REPLA(y$, Chr$(96), 92, 1)
                        Call GRAREG("MASTER", y$, REGMA&)
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
            x$ = REGLEIDO$("MASTER", II& + 1)
            If CVI(Left$(x$, 2)) <> II& Then
                Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                Screen.MousePointer = 11
                CIX% = II&
                Call REPLA(x$, MKI$(CIX%), 1, 2)
                Call GRAREG("MASTER", x$, II& + 1)
            End If
            COD$ = AJUSTI$(TRIM$(Mid$(x$, 77, 15)), 15) + Mid$(x$, 92, 1)
            If Left$(COD$, 4) = "zzz-" Then
              If XVALO(Mid$(COD$, 5, 6)) > RECUMAX& Then
                RECUMAX& = XVALO(Mid$(COD$, 5, 6))
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
        x$ = REGBLAN$("INVERT")
        Call REPLA(x$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", x$, 1)
        '
        J& = 1
        FIN& = NV% - 1
        For II& = 0 To FORMALTER.LINOSORT.ListCount - 1
            Call TRACE(20, II&, FIN&)
            FORMALTER.LINOSORT.ListIndex = II&
            CODISEL$(II&) = Mid$(FORMALTER.LINOSORT.TEXT, 2, 16)
            CISEL%(II&) = XVALO(Mid$(FORMALTER.LINOSORT.TEXT, 18))
            x$ = REGBLAN$("INVERT")
            Call REPLA(x$, CODISEL$(II&), 1, 16)
            Call REPLA(x$, MKI$(CISEL%(II&)), 17, 2)
            Call GRAREG("INVERT", x$, II& + 2)
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
       For J& = ULIS& - 1 To 0 Step -1
         If Left$(FORMALTER.LINOSORT.List(J&), 17) <= Left$(RAGRE$, 17) Then
           FORMALTER.LINOSORT.AddItem RAGRE$, J& + 1
           GoTo 10
         End If
       Next J&
       FORMALTER.LINOSORT.AddItem RAGRE$, 0
10 End If
   '
End Sub

Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta Definir Depósitos")
        OPMAST07.Show 1
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
          x$ = REGLEIDO$("TATIMAT", U&)
          y$ = REGBLAN$("TABTIMA")
          Call REPLA(y$, x$, 1, 16)
          Call REPLA(y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", y$)
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
   'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   On Error Resume Next
   SECOPTEMP.MoveFirst
     If Err Then
       On Error GoTo 0
       Screen.MousePointer = 1
       Call COMUNI("Copia Imposible.\Código Origen sin Operaciones.")
       GoTo 10
     End If
     On Error GoTo 0
   Set SECOPTEMP = Nothing
     '
   SQLX$ = "SELECT * FROM Secoper"
   SQLX$ = SQLX$ & " WHERE CodIConj = " & CODEST%
   'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   SECOPTEMP.MoveFirst
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
   Set SECOPTEMP = Nothing

   '
End Sub
'
Sub LISTRUTAS()
   '
   CI1% = PRICOD%
   CI2% = ULTCOD%
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(x$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Dim PUNCOD$(32767)
   '
   CI1% = CVI(Left$(x$, 2))
   COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(x$, 3, 2))
   COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
   '
   'CUENTO LA CANTIDAD DE REGISTROS SEGUN CONDICION
   SQLX$ = "SELECT COUNT (*)AS CANTI FROM Master "
   SQLX$ = SQLX$ + "WHERE CODIGO >='" & COD1$ & "' "
   SQLX$ = SQLX$ + "AND CODIGO <='" & COD2$ & "' "
   Dim CANTTTREG As ADODB.Recordset
   Set CANTTTREG = FLEXCONN.Execute(FILTSQL$(SQLX$))

   SQLX$ = "SELECT * FROM Master "
   SQLX$ = SQLX$ + "WHERE CODIGO >='" & COD1$ & "' "
   SQLX$ = SQLX$ + "AND CODIGO <='" & COD2$ & "' "
   SQLX$ = SQLX$ + "ORDER BY CODIGO "
   'Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   Dim MasterTemp As ADODB.Recordset
   Set MasterTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   On Error Resume Next
   MasterTemp.MoveFirst
   If Err > 0 Then Exit Sub
   '
   'CAREC% = MasterTemp.RecordCount
   CAREC% = CANTTTREG!CANTI
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
     SQLX$ = "SELECT * FROM Secoper WITH(NOLOCK)"
     SQLX$ = SQLX$ & " WHERE CodIConj = " & CI0%
     SQLX$ = SQLX$ + "ORDER BY NumeOper, AltOper ASC"

     'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim SECOPTEMP As ADODB.Recordset
     Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With SECOPTEMP
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
       y$ = REGBLAN$("SECOPSEL")
       Call REPLA(y$, MKI$(CI0%), 1, 2)
       Call REPLA(y$, DESCRI$, 3, 60)
       V& = V& + 1
       Call GRAREG("SECOPSEL", y$, V&)
       '
       Call REPLA(y$, Space$(60), 3, 60)
       V& = V& + 1
       Call GRAREG("SECOPSEL", y$, V&)
       '
       Do Until .EOF = True
         NOPY% = !NumeOper
         ALTY% = !AltOper
         DOPE$ = !DescOper
         LOSTAND = LOTESTAN(CI0%)
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
         CAOPS2 = !CantiOps2
         CATOP% = !CategOps
         CATOP2% = !CategOps2
         EQUIPO$ = !CodMaq0
         EQUIP1$ = !CodMaq1
         EQUIP2$ = !CodMaq2
         EQUIP3$ = !CodMaq3
         EQUIP4$ = !CodMaq4
         EQUIP5$ = !CodMaq5
         '
         PROTERI% = !ProTerOp
         PRECTER# = !PreUnid
         MONETER% = !MonePrec
         FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
         '
         NOPES$ = TRIM$(FORMATNUM$(NOPY%, "I4"))
         While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
         If ALTY% > 0 Then
             NOPES$ = NOPES$ + "-" + TRIM$(Str$(ALTY%))
         End If
         DOPES$ = DOPE$
         TVAR = MINVARI * LOSTAND
         '
         Call REPLA(y$, NOPES$, 3, 8)
         Call REPLA(y$, DOPES$, 11, 46)
         Call REPLA(y$, EQUIPO$, 57, 6)
         Call REPLA(y$, MKS$(HORPREP), 63, 4)
         Call REPLA(y$, MKS$(HORFIJO), 67, 4)
         Call REPLA(y$, MKS$(TVAR), 71, 4)
         Call REPLA(y$, MKS$(DEMODIS), 75, 4)
         Call REPLA(y$, MKS$(HORLIMP), 79, 4)
         Call REPLA(y$, MKS$(CAOPS), 83, 4)
         Call REPLA(y$, MKS$(LOSTAN), 91, 4)
         Call REPLA(y$, EQUIP1$, 103, 6)
         Call REPLA(y$, EQUIP2$, 109, 6)
         Call REPLA(y$, EQUIP3$, 115, 6)
         '
         TSTANBA = MINVARI    ' TIESTAN(!COD_INTE, !NumeOper, 1, 1, ALTOP%)
         CADEN = 0
         If TSTANBA > 0 Then CADEN = 60 / TSTANBA
         Call REPLA(y$, MKS$(CADEN), 125, 4)
         '
         V& = V& + 1
         Call GRAREG("SECOPSEL", y$, V&)
         .MoveNext
       Loop
       '
     End With
     Set SECOPTEMP = Nothing
     '
     Call REPLA(y$, String$(126, 0), 3, 126)
     V& = V& + 1
     Call GRAREG("SECOPSEL", y$, V&)
     '
     y$ = REGBLAN$("SECOPSEL")
     V& = V& + 1
     Call GRAREG("SECOPSEL", y$, V&)
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
    Dim CIANU%(32767) '\\\OT-06-0078-WAR-03/03/06///
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
      If CAIT% > 0 Then CIANU%(CIBAS%) = 1
      For J% = 1 To CAIT%
        CIANU%(CIJ%(J%)) = 1 '\\\OT-06-0078-WAR-03/03/06/// - BANDERA QUE E COMPONENTE DE ALGUNA ESTRUCTURA
        For K% = J% + 1 To CAIT%
          If CIJ%(J%) = CIJ%(K%) Then
            COD1$ = TRIM$(CODEXT$(CIBAS%))
            COD2$ = TRIM$(CODEXT$(CIJ%(J%)))
            If CODINT%(COD1$) = CIBAS% Then ' SUPRIME ITEMS DE BAJA
              CAERR% = CAERR% + 1
              MENSA$ = "Repetición del Código " + COD2$ + "\en Fórmula de Codigo " + COD1$
              If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
            End If
          End If
        Next K%
      Next J%
      '
      Call EXPLOMUL(CIBAS%)
      For J% = 2 To CONTAMUL%
        If CIK%(J%) = CIBAS% Then
          CAERR% = CAERR% + 1
          MENSA$ = "Recursividad en Código " + TRIM$(CODEXT$(CIBAS%)) + "\" + DESCRIT0$(CIBAS%)
          If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
        End If
      Next J%
      '
89  Next I&
    '
    If CAERR% < 1 Then
        Call COMUNI("Control Recursividad O.K. - No se Detectaron Repeticiones")
      Else
        Call MENSERR(20, "Se Encontraron" + Str$(CAERR%) + " Repeticiones en Despieces")
    End If
    '
    '\\\OT-06-0078-WAR-03/03/06///
    NMS% = NUMAS%
    HCIB% = 0
    For J% = 1 To NMS%
      If CIANU%(J%) = 1 Then
        CDXX$ = CODEXT$(J%)
        If CODINT%(CDXX$) < 0 Then ' SIGNIFICA QUE ESTA DE BAJA
          CIANU%(J%) = 2
          HCIB% = 1 'BANDERA PARA SABER SI EXISTE ALGUN CI DE BAJA EN ALGUNA ESTRUCTURA
        End If
      End If
    Next J%
    If HCIB% = 1 Then
      RESP1% = COMALTER%("Existen Conjuntos o Componentes de Baja\en Alguna(s) Fórmula(s) de Conjunto.\  \Puede Ahora Emitir Listado de Control.\\Listar\Continuar")
      If RESP1% = 2 Then GoTo 99
      y$ = REGBLAN$("COMPOBAJ")
      JJ& = 0
      For J% = 1 To NMS%
        If CIANU%(J%) = 2 Then
          Call REPLA(y$, MKI$(J%), 1, 2)
          JJ& = JJ& + 1
          Call GRAREG("COMPOBAJ", y$, JJ&)
        End If
      Next J%
      Call SETULTREG("COMPOBAJ", JJ&)
      Call FINAL("*COMPOBAJ")
    Else
      Call COMUNI("No se Encontraron Componentes de Baja\en Ninguna Fórmula de Conjunto.")
    End If
    '\\\OT-06-0078-WAR-FIN///
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
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
        Exit Sub
    End If
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    FIECO% = FILENBR%("LIESTRU")
    Dim ECM As String * 16
    '
    Z00$ = REGBLAN$("LIESTRU")
    Y00$ = REGBLAN$("LISTRULA")
    V00$ = REGBLAN$("LISTRUAT")
    '
    Call HEADERS("LIESTRU", "")
    Call HEADERS("LIESTRU", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LIESTRU", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '\\\OT-05-0702-WAR-02/11/05///
    Call HEADERS("LISTRULA", "")
    Call HEADERS("LISTRULA", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LISTRULA", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '
    Call HEADERS("LISTRUAT", "")
    Call HEADERS("LISTRUAT", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LISTRUAT", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '\\\OT-05-0702-WAR-02/11/05///
    '
    Screen.MousePointer = 11
    FIN& = NUMAS%
    J& = 0
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
                    UM$ = UNIMED$(CI1%)
                    Call CONVERUNID(CAN, UM$)
                    If Abs(CAN) > 0.00005 Then
                      Z$ = Z00$
                      y$ = Y00$
                      '\\\OT-05-0702-WAR-01/11/05///
                      V$ = V00$
                      '\\\OT-05-0702-WAR-FIN///
                      Call REPLA(Z$, MKI$(CI%), 1, 2)
                      Call REPLA(Z$, MKI$(CI1%), 3, 2)
                      Call REPLA(Z$, UM$, 5, 2)
                      Call REPLA(Z$, MKS$(CAN), 7, 4)
                      Call REPLA(Z$, Chr$(PD%), 11, 1)
                      '\\\OT-06-0352-MH-30/08/06///
                      If STATITEM%(CI1%) = -1 Then
                         Call REPLA(Z$, "Baja", 13, 4)
                      End If
                      '\\\OT-06-0352-MH-FIN///
                      '
                      '\\\OT-05-0702-WAR-01/11/05///
                      Call REPLA(V$, MKI$(CI%), 1, 2)
                      Call REPLA(V$, MKI$(CI1%), 3, 2)
                         ALI1$ = NUPLANO$(CI1%)
                      Call REPLA(V$, ALI1$, 5, 24)
                         ALI1$ = ""
                      Call REPLA(V$, UM$, 29, 2)
                         ATR1$ = ATRIB1$(CI1%)
                      Call REPLA(V$, ATR1$, 31, 32)
                         ATR2$ = ATRIB2$(CI1%)
                      Call REPLA(V$, ATR2$, 63, 32)
                      Call REPLA(V$, MKS$(CAN), 95, 4)
                      Call REPLA(V$, Chr$(PD%), 99, 1)
                         REFE$ = Mid$(EXPLOLIN$(K%), 129)
                      Call REPLA(V$, REFE$, 109, 64)
                      '\\\OT-06-0352-MH-30/08/06///
                      If STATITEM%(CI1%) = -1 Then
                         Call REPLA(V$, "Baja", 173, 4)
                      End If
                      '\\\OT-06-0352-MH-FIN///
                      '\\\OT-05-0702-WAR-FIN///
                      '
                      Call REPLA(y$, MKI$(CI%), 1, 2)
                      Call REPLA(y$, MKI$(CI1%), 3, 2)
                      Call REPLA(y$, UM$, 5, 2)
                        PESOMET = PESMETRO(CI1%)
                      Call REPLA(y$, MKS$(PESOMET), 7, 4)
                        'LARGO CORTE (MM)
                        LACORT = XVALO(Mid$(EXPLOLIN$(K%), 111, 6))
                      Call REPLA(y$, MKS$(LACORT), 11, 4)
                      Call REPLA(y$, MKS$(CAN), 19, 4)
                        'CANTIDAD
                        CANTIUN = XVALO(Mid$(EXPLOLIN$(K%), 89, 9))
                      Call REPLA(y$, MKS$(CANTIUN), 15, 4)
                      Call REPLA(y$, Chr$(PD%), 11, 1)
                        'REFERENCIA
                        REFE$ = Mid$(EXPLOLIN$(K%), 129)
                      Call REPLA(y$, REFE$, 33, 64)
                      '\\\OT-06-0352-MH-30/08/06///
                      If STATITEM%(CI1%) = -1 Then
                         Call REPLA(y$, "Baja", 97, 4)
                      End If
                      '\\\OT-06-0352-MH-FIN///
                      
                      '
                      J& = J& + 1
                      Call GRAREG("LIESTRU", Z$, J&)
                      Call GRAREG("LISTRULA", y$, J&)
                      '\\\OT-05-0702-WAR-01/11/05///
                      Call GRAREG("LISTRUAT", V$, J&)
                      '\\\OT-05-0702-WAR-FIN///
                      ACTIX% = 1
                      '\\\OT-05-0663-WAR-FIN///
                    End If
                  End If
                End If
              Next K%
            End If
          End If
        End If
        '
        If ACTIX% > 0 Then
          Z$ = Z00$
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          J& = J& + 1
          Call GRAREG("LIESTRU", Z$, J&)
          '\\\OT-05-0663-WAR-07-10-05///
          y$ = Y00$
          Call REPLA(y$, MKI$(CI%), 1, 2)
          Call GRAREG("LISTRULA", y$, J&)
          '\\\OT-05-0663-WAR-FIN///
          '\\\OT-05-0702-WAR-01/11/05///
          V$ = V00$
          Call REPLA(V$, MKI$(CI%), 1, 2)
          Call GRAREG("LISTRUAT", V$, J&)
          '\\\OT-05-0702-WAR-FIN///
        End If
        '
      End If
      '
    Next I&
    '
    Screen.MousePointer = 1
    '\\\OT-05-0663-WAR-07-10-05///
      RTA% = COMALTER%("Elija Formato de Salida:\\Compacto\Detallado\Atributos")
      If RTA% = 1 Then
          Call SETULTREG("LIESTRU", J&)
          Call CIERRARCH("LIESTRU")
          Call FILESORT("LIESTRU", "", 1, 1, "ASC")
          '
          Call FINAL("*LIESTRU")
        ElseIf RTA% = 2 Then
          '\\\OT-05-0663-WAR-07-10-05///
          Call SETULTREG("LISTRULA", J&)
          Call CIERRARCH("LISTRULA")
          Call FILESORT("LISTRULA", "", 1, 1, "ASC")
          '\\\OT-05-0663-WAR-FIN///
          Call FINAL("*LISTRULA")
        ElseIf RTA% = 3 Then
          '\\\OT-05-0702-WAR-01/11/05///
          Call SETULTREG("LISTRUAT", J&)
          Call CIERRARCH("LISTRUAT")
          Call FILESORT("LISTRUAT", "", 1, 1, "ASC")
          '\\\OT-05-0702-WAR-FIN///
          Call FINAL("*LISTRUAT")
      End If
    '\\\OT-05-0663-WAR-FIN///
    '
End Sub
'
Sub LIMPLOSI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
60 Screen.MousePointer = 11
   IMPLOSI07.List1.Clear
   IMPLOSI07.Caption = "Implosión " + DESCRIT$(CI%)
   IMPLOSI07.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI07.Left = Left + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left + (Width - ScaleWidth) / 2
   'IMPLOSI07.Show
   '
   'Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas WHERE FORMULAS!CODIELEM=" & CI%, dbOpenDynaset, dbReadOnly)
   Dim FORMULASTEMP As ADODB.Recordset
   Set FORMULASTEMP = New ADODB.Recordset
   FORMULASTEMP.CursorType = adOpenKeyset
   FORMULASTEMP.LockType = adLockReadOnly
   SQLX$ = "SELECT * FROM Formulas WHERE CODIELEM=" & CI%
   FORMULASTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
   
   On Error Resume Next
   FORMULASTEMP.MoveLast
   If Err Then GoTo 99
   CAREC% = FORMULASTEMP.RecordCount
   If CAREC% < 1 Then GoTo 99
   '
   FORMULASTEMP.MoveFirst
   ACTIREC% = 1
61 'If FORMULASTEMP.NoMatch = False Then
    If Not (FORMULASTEMP.EOF And FORMULASTEMP.BOF) Then 'si no esta vacio por lo tanto lo encuentra
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
          IMPLOSI07.List1.AddItem Z$
          On Error Resume Next
          IMPLOSI07.List1.TopIndex = IMPLOSI07.List1.ListCount - 10
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
   If PORMILLAR > 0 Then
     IMPLOSI07.Caption = IMPLOSI07.Caption + " - POR MILLAR"
   End If
   '
99 Screen.MousePointer = 1
   '
   FORMULASTEMP.Close
   Set FORMULASTEMP = Nothing
   
   IMPLOSI07.Show 1
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
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    REGQ& = 0:
    REGR& = 0
    '
    OPX% = COMALTER%("Formato de Salida:\\Normal\Con Referencias")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim TIESTRU%(32767), ESPARSTRU%(32767)
    '
    NMS% = NUMAS%
    FIN& = NUMAS%
    J& = 0
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
    'Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
    Dim MasterTemp As ADODB.Recordset
    Set MasterTemp = New ADODB.Recordset
    MasterTemp.CursorType = adOpenKeyset
    MasterTemp.LockType = adLockReadOnly
    MasterTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
      
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
            If OPX% = 2 Then
               REGR& = REGR& + 1
               Call GRAREG("LIMULREF", String$(256, "="), REGR&)
            End If
          End If
          '
          For II% = 1 To CONTAMUL%
            '
            If II% = 1 Then
              XX$ = REGBLAN$("LIMULTI")
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
              If OPX% = 2 Then
                 REGR& = REGR& + 1
                 YY$ = REGBLAN$("LIMULREF")
                 Call GRAREG("LIMULREF", YY$, REGR&)
              End If
            End If
            '
            CANTU = CANT(II%)
            UME$ = UNIMED$(CIK%(II%))
            Call CONVERUNID(CANTU, UME$)
            '
            If CANTU > 0.00005 Then  ' SUPRIME DESARMADO
              XX$ = REGBLAN$("LIMULTI")
              Call REPLA(XX$, MKI$(II% - 1), 1, 2)
              Call REPLA(XX$, CODEXT$(CIK%(II%)), 3, 16)
              Call REPLA(XX$, DESCRIT$(CIK%(II%)), 17 + 2 * NIVE%(II%), 48)
              Call REPLA(XX$, " " + UME$, 69, 5)
              Call REPLA(XX$, CSTRING$(MKS$(CANTU), 4, 10, 5, 2), 75, 10)
              Call REPLA(XX$, CSTRING$(MKI$(PAD%(II%)), 1, 6, 0, 2), 85, 6)
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
              '
              If OPX% = 2 Then
                Call FRATEXTO(REFEX$(II%), 60)
                If CARETEX% < 1 Then CARETEX% = 1
                For JKI% = 1 To CARETEX%
                   YY$ = REGBLAN$("LIMULREF")
                   Call REPLA(YY$, XX$, 1, 96)
                   Call REPLA(YY$, RETEX$(JKI%), 97, 60)
                   REGR& = REGR& + 1
                   Call GRAREG("LIMULREF", YY$, REGR&)
                Next JKI%
              End If
            End If
            '
            If II% = 1 Then
              XX$ = REGBLAN$("LIMULTI")
              REGQ& = REGQ& + 1
              Call GRAREG("LIMULTI", XX$, REGQ&)
              '
              If OPX% = 2 Then
                 REGR& = REGR& + 1
                 YY$ = REGBLAN$("LIMULREF")
                 Call GRAREG("LIMULREF", YY$, REGR&)
              End If
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
    Call HEADERS("LIMULTI", "")
    If OPX% = 2 Then
      Call SETULTREG("LIMULREF", REGR&)
      Call HEADERS("LIMULREF", "")
    End If
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    If OPX% = 1 Then
        Call FINAL("*LIMULTI")
      Else
        Call FINAL("*LIMULREF")
    End If
    '
    MasterTemp.Close
    Set MasterTemp = Nothing
End Sub
'

Sub LIMULIN()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
       Exit Sub
    End If
    '
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    J& = 0
    'Set MasterTemp = Articulos.OpenRecordset("SELECT * FROM Master WHERE CODIGO>='" & COD1$ & "' AND CODIGO<='" & COD2$ & "' ORDER BY CODIGO ASC;", dbOpenDynaset, dbReadOnly)
    Dim MasterTemp As ADODB.Recordset
    Set MasterTemp = New ADODB.Recordset
    MasterTemp.CursorType = adOpenKeyset
    MasterTemp.LockType = adLockReadOnly
    SQLX$ = ("SELECT * FROM Master WHERE CODIGO>='" & COD1$ & "' AND CODIGO<='" & COD2$ & "' ORDER BY CODIGO ASC")
    MasterTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
  
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
              UM$ = UNIMED$(CIX1%)
              Call CONVERUNID(CAN, UM$)
              If Abs(CAN) > 0.0005 Then
                Z$ = REGBLAN$("LIESTRU")
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                Call REPLA(Z$, MKI$(CIX1%), 3, 2)
                Call REPLA(Z$, UM$, 5, 2)
                Call REPLA(Z$, MKS$(CAN), 7, 4)
                Call REPLA(Z$, Chr$(PD%), 11, 1)
                '
                J& = J& + 1
                Call GRAREG("LIESTRU", Z$, J&)
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
          J& = J& + 1
          Call GRAREG("LIESTRU", Z$, J&)
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
    Call SETULTREG("LIESTRU", J&)
    Call CIERRARCH("LIESTRU")
    '
    Call HEADERS("LIESTRU", "")
    Screen.MousePointer = 1
    Call FINAL("*LICOMPRA")
    '
    MasterTemp.Close
    Set MasterTemp = Nothing
    
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
     x$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(x$, 77, 15)
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
     COU = XVALO(COSUNIX$)
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
19 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
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
     If TRIM$(CODPAR1$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 143, 6))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = XVALO(ENTE$) + (XVALO(DECI$)) / 100000
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
       x$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(x$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
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
     x$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(x$, 77, 15)
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
     COU = XVALO(COSUN$)
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
19 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
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
         USO = XVALO(ENTE$) + (XVALO(DECI$)) / 100000
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
       x$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(x$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
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
   Label79 = ""
   Label74 = ""
   Label114 = ""
   Label116 = ""
   DoEvents
   '
End Sub

Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   On Error Resume Next
   Command26.SetFocus
   On Error GoTo 0
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
   IMPLOSI07.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI07.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI07.List1.Clear
   IMPLOSI07.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   IMPLOSI07.Show
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
     IMPLOSI07.List1.AddItem Z$
   Next KKI%
   Screen.MousePointer = 1
   '
End Sub

Sub ACVEPROINT()
   For UI& = 0 To 32767
     PROVINT%(UI&) = 0
   Next UI&
   '
   For U& = 1 To ULTREG&("TAPROINT")
      x$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(x$, 2))
      If NP% > 0 Then
        PROVINT%(NP%) = 1
      End If
   Next U&
   Call CIERRARCH("TAPROINT")
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
          TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 55, 10)
          TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 204, 10)
          TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 65, 10)
          TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 75, 10)
          TETABU1$(9, CAITAB1%) = Mid$(TTXX$, 85, 10)
          SUMAPAR# = SUMAPAR# + XVALO(Mid$(TTXX$, 85, 10))
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

'\\\OT-05-0633-WAR-22/09/05///
Sub CALCUKILOS(CI1%)
  '
  KG = PESMETRO(CI1%)
  LM = XVALO(Text42)
  CANTI1 = XVALO(Text41)
  If KG > 0.005 And LM > 0.005 Then
    KGU = LM * KG / 1000
    Label36 = FORMATNUM$(KGU, "F8.4") 'KG/UNIDAD
    USU1 = KGU * CANTI1
    Text32 = FORMATNUM$(USU1, "F8.4") 'USO ES IGUAL A FORMULA (KG/UNIDAD)* CANTIDAD
  ElseIf CANTI1 > 0 Then
    Label36 = "" 'KG/UNIDAD=0
    Text32 = FORMATNUM$(CANTI1, "F8.4") 'USO ES IGUAL A CANTIDAD
  End If
  '
End Sub
'\\\OT-05-0633-WAR-FIN///
'\\\OT-06-0252-WAR-22/06/06///
Function MSTMU#(COTMU$)
  FIN& = ULTREG&("TABLATMU") 'Base de Datos de TMU's
  For J& = 1 To FIN&
    x$ = REGLEIDO$("TABLATMU", J&)
    If COTMU$ = Mid$(x$, 1, 4) Then
      MSTMU# = CVD(Mid$(x$, 37, 8)) 'Duración en mseg.
      Exit Function
    End If
  Next J&
  MSTMU# = 0
End Function
Sub LISTMUS()
   '
   Screen.MousePointer = 11
   '
   CI1% = PRICOD%
   CI2% = ULTCOD%
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(x$) < 4 Then
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   CI1% = CVI(Left$(x$, 2))
   COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(x$, 3, 2))
   COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
   '
   Call ABREMASTER
   Call ABRESECOP
   SQLX$ = "SELECT * FROM TMUSOPER "
   SQLX$ = SQLX$ + "ORDER BY COD_OPER ASC, NUME_ORD ASC "
   'Set TMUTMP = Articulos.OpenRecordset(SQLX$, 4)
   Dim TMUTMP As ADODB.Recordset
   Set TMUTMP = New ADODB.Recordset
   TMUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With TMUTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "No Existen TMU's que\Puedan Listarse")
      GoTo 99
    End If
    On Error GoTo 0
    JJ& = 0
    ZBLA$ = REGBLAN$("LISTMUS")
    Do While .EOF = False
      ZZ$ = ZBLA$
      CIXI% = !COD_INTE
      If CODEXT$(CIXI%) >= COD1$ Then
        If CODEXT$(CIXI%) <= COD2$ Then
          NOPE% = !Cod_Oper
          DOPE1$ = DESCRIOP$(CIXI%, NOPE%)
          TMU1$ = !Cod_TMU
          DURA1# = MSTMU#(TMU1$)
          CANTI1# = !Cantidad
          TOTMS1# = DURA1# * CANTI1#
          '
          Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
          Call REPLA(ZZ$, MKI$(NOPE%), 3, 2)
          Call REPLA(ZZ$, DOPE1$, 5, 23)
          Call REPLA(ZZ$, TMU1$, 28, 4)
          Call REPLA(ZZ$, MKD$(DURA1#), 32, 8)
          Call REPLA(ZZ$, MKD$(CANTI1#), 40, 8)
          Call REPLA(ZZ$, MKD$(TOTMS1#), 48, 8)
          JJ& = JJ& + 1
          Call GRAREG("LISTMUS", ZZ$, JJ&)
        End If
      End If
      .MoveNext
    Loop
    Call SETULTREG("LISTMUS", JJ&)
    Call FILESORT("LISTMUS", "LISTMUS", 1, 1, "ASC")
    Call HEADERS("LISTMUS", "")
    Call HEADERS("LISTMUS", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LISTMUS", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    Call FINAL("*LISTMUS")
    '
   End With
   TMUTMP.Close
   Set TMUTMP = Nothing
99 Screen.MousePointer = 1
   '
End Sub
'\\\OT-06-0252-WAR-FIN///

Sub ACTECOMAS()
   '
   Call GENECOMAS 'RUTINA ACUTALIZADA CON CARGALISEL
   Exit Sub
   
   
   
   If UPDATMASTER% = 1 Then
      EL2$ = REGBLAN$("ECOMAST")
      LOFI1% = LENFIELD%("ECOMAST", 1)
      POFI2% = POSFIELD%("ECOMAST", 2)
      POFI3% = POSFIELD%("ECOMAST", 3)
      Call REPLA(EL2$, Text1, 1, LOFI1%)
        DEPLANO$ = TRIM$(Text3)
        NXPLANO$ = TRIM$(Text45)
        If NXPLANO$ <> "" Then
          'If NXPLANO$ <> TRIM$(TRIM$(Text1)) Then
          If InStr(TRIM$(Text1), NXPLANO$) < 1 Then
            While Len(DEPLANO$ + NXPLANO$) > 43
              DEPLANO$ = Left$(DEPLANO$, Len(DEPLANO$) - 1)
            Wend
            DEPLANO$ = DEPLANO$ + " (" + NXPLANO$ + ")"
          End If
        End If
      Call REPLA(EL2$, DEPLANO$, POFI2%, 46)
      Call REPLA(EL2$, MKI$(XVALO(Label28)), POFI3%, 2)
      Call FILTERPUTRECORD("ECOMAST", 1, Left$(EL2$, LOFI1%), EL2$)
      Call FILESORT("ECOMAST", "", 1, 1)
      Call FRESHECHO("ECOMAST")
   End If
   '
   UPDATMASTER% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub MUESTRAPDF(ARFULLNAME$)
    '
    IDME$ = TRIM$(Mid$(List3.TEXT, 65, 4))
    '
    On Error Resume Next
    SHOWPDF.List1.Clear
    SHOWPDF.Label1.Caption = ARFULLNAME$ + AJUSTI$(USERID$, 16)
    SHOWPDF.Pdf1.LOADFILE$ (ARFULLNAME$)
    If Err Then
         On Error GoTo 0
         Call MUESTRADOC(ARFULLNAME$)
      Else
         On Error GoTo 0
         SHOWPDF.Label4 = ARNAME$
         SHOWPDF.Label5 = IDME$
         If IDME$ <> "" Then
            For JJ1& = 1 To List3.ListCount
               IDME1$ = TRIM$(Mid$(List3.List(JJ1& - 1), 65, 4))
               If IDME1$ = IDME$ Then
                  If AJUSTI$(List3.List(JJ1& - 1), 48) <> ARFULLNAME$ Then
                     SHOWPDF.List1.AddItem AJUSTI$(List3.List(JJ1& - 1), 48)
                  End If
               End If
            Next JJ1&
         End If
         SHOWPDF.Show 1
    End If
    '
End Sub
'
Sub MUESTRADOC(ARFULLNAME$)
    '
    ' Ejecutar un acceso directo
    Call ShellExecute(Me.hwnd, "Open", ARFULLNAME$, "", "", 1)
    '
End Sub
'
Sub RecuStoMinLotEco()
   '
   Call ABRECONEXION
   '
   QTR& = CantRegistros("MASTER", "STOMIN < 0 OR LOREPOS < 0")
   If QTR& < 1 Then Exit Sub
   '
   Call COMUNI("FLEXOFT Debe Ahora Reorganizar\el Maestro de Artículos.\   \Esto Puede Demandar Algunos Minutos.")
   Screen.MousePointer = 11
   'Recupera Stock Minimo
   SQLX$ = "SELECT * FROM MASTER WITH(NOLOCK) WHERE STOMIN < 0"
   Dim MAST As ADODB.Recordset
   Set MAST = New ADODB.Recordset
   Set MAST = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With MAST
     Do While Not .EOF
        
        STOMINIMO = Abs(XVALO(!STOMIN)) * 100
        CONDI$ = "Codint  = " & XVALO(!CODINT)
        Call ACTUREGISTRO("MASTER", "Stomin", CONDI$, STOMINIMO, REG&)
        .MoveNext
     Loop
     .Close
   End With
   Set MAST = Nothing
   '
   '
   'Recupera Lote Economico
   SQLX$ = "SELECT * FROM MASTER WITH(NOLOCK) WHERE Lorepos < 0"
   Set MAST = New ADODB.Recordset
   Set MAST = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With MAST
     Do While Not .EOF
        Loreposicion = Abs(XVALO(!LOREPOS)) * 100
        CONDI$ = "Codint  = " & XVALO(!CODINT)
        ' Call ACTUREGISTRO("MASTER", "Stomin", CONDI$, Loreposicion, REG&)
        Call ACTUREGISTRO("MASTER", "Lorepos", CONDI$, Loreposicion, REG&)
        .MoveNext
     Loop
     .Close
   End With
   Set MAST = Nothing
   Screen.MousePointer = 1
   '
End Sub

