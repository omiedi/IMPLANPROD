VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form AMASTOAHP 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos - A-B-M-L Maestro de Items de Stock"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13785
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6600
      ScaleHeight     =   75
      ScaleWidth      =   3345
      TabIndex        =   391
      Top             =   1150
      Width           =   3405
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   392
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Index           =   1
      Left            =   12120
      TabIndex        =   332
      Top             =   120
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   265
      Left            =   7440
      TabIndex        =   185
      Top             =   8520
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Frame Frame16 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   265
      Left            =   120
      TabIndex        =   184
      Top             =   8520
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.TextBox Text38 
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
      Left            =   1275
      TabIndex        =   182
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
      TabIndex        =   180
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   3820
      Picture         =   "AMASTOAHP.frx":0000
      TabIndex        =   3
      Top             =   240
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Height          =   600
      Left            =   1275
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   630
      Width           =   5295
   End
   Begin VB.TextBox Text45 
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
      Left            =   4725
      MaxLength       =   25
      TabIndex        =   89
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
      TabIndex        =   117
      ToolTipText     =   "Tabla de Marcas y Modelos"
      Top             =   220
      Visible         =   0   'False
      Width           =   260
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00E2D3C0&
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
      Index           =   20
      Left            =   10545
      TabIndex        =   84
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
         Picture         =   "AMASTOAHP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   86
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
         Picture         =   "AMASTOAHP.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   8310
      Left            =   14520
      ScaleHeight     =   8250
      ScaleWidth      =   1005
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
      Begin VB.CommandButton Command58 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   174
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
         Picture         =   "AMASTOAHP.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Actualizar Precios de Venta"
         Top             =   4200
         Width           =   650
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
         Picture         =   "AMASTOAHP.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   87
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
         Picture         =   "AMASTOAHP.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   83
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
         Picture         =   "AMASTOAHP.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   82
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
         Picture         =   "AMASTOAHP.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   81
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
         Picture         =   "AMASTOAHP.frx":1150
         Style           =   1  'Graphical
         TabIndex        =   80
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
         Picture         =   "AMASTOAHP.frx":145A
         Style           =   1  'Graphical
         TabIndex        =   79
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
         Picture         =   "AMASTOAHP.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   78
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
         Picture         =   "AMASTOAHP.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   77
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
         Picture         =   "AMASTOAHP.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   76
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
         Picture         =   "AMASTOAHP.frx":251A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "AMASTOAHP.frx":2824
         Top             =   7455
         Width           =   1515
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   4
      Top             =   240
      Width           =   2565
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2D3C0&
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
      Index           =   19
      Left            =   12180
      TabIndex        =   19
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
         Picture         =   "AMASTOAHP.frx":4746
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "AMASTOAHP.frx":4A50
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   630
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   14865344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Insumos"
      TabPicture(0)   =   "AMASTOAHP.frx":4D5A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MARCOFICHA(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Componentes"
      TabPicture(1)   =   "AMASTOAHP.frx":4D76
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Elaboración"
      TabPicture(2)   =   "AMASTOAHP.frx":4D92
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Especificación"
      TabPicture(3)   =   "AMASTOAHP.frx":4DAE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MARCOFICHA(1)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Stocks"
      TabPicture(4)   =   "AMASTOAHP.frx":4DCA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label15(2)"
      Tab(4).Control(1)=   "Frame15(5)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Histórico de Movimientos"
      TabPicture(5)   =   "AMASTOAHP.frx":4DE6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label15(8)"
      Tab(5).Control(1)=   "Frame15(20)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Estadísticas"
      TabPicture(6)   =   "AMASTOAHP.frx":4E02
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame15(21)"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame15 
         Height          =   6480
         Index           =   21
         Left            =   -74880
         TabIndex        =   381
         Top             =   360
         Width           =   13335
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2685
            Index           =   5
            Left            =   60
            TabIndex        =   382
            Top             =   360
            Width           =   10150
            _ExtentX        =   17912
            _ExtentY        =   4736
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2685
            Index           =   6
            Left            =   10200
            TabIndex        =   383
            Top             =   360
            Width           =   3100
            _ExtentX        =   5477
            _ExtentY        =   4736
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2925
            Index           =   7
            Left            =   60
            TabIndex        =   384
            Top             =   3360
            Width           =   10150
            _ExtentX        =   17912
            _ExtentY        =   5159
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2925
            Index           =   8
            Left            =   10200
            TabIndex        =   385
            Top             =   3360
            Width           =   3100
            _ExtentX        =   5477
            _ExtentY        =   5159
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   8520
            TabIndex        =   390
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Últimas  Cotizaciones"
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
            Index           =   63
            Left            =   0
            TabIndex        =   389
            Top             =   120
            Width           =   10155
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Últimas Ventas"
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
            Index           =   64
            Left            =   0
            TabIndex        =   388
            Top             =   3120
            Width           =   10155
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Estadísticas de Cotizaciones"
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
            Index           =   65
            Left            =   10200
            TabIndex        =   387
            Top             =   120
            Width           =   3195
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Estadisticas de Ventas"
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
            Index           =   66
            Left            =   10200
            TabIndex        =   386
            Top             =   3120
            Width           =   3195
         End
      End
      Begin VB.Frame Frame15 
         Height          =   6495
         Index           =   20
         Left            =   -74880
         TabIndex        =   370
         Top             =   360
         Width           =   13335
         Begin VB.CommandButton Command40 
            Caption         =   "..."
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
            Index           =   7
            Left            =   13065
            Picture         =   "AMASTOAHP.frx":4E1E
            TabIndex        =   375
            Top             =   240
            Width           =   175
         End
         Begin VB.TextBox Text6 
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
            Index           =   1
            Left            =   10185
            TabIndex        =   374
            Text            =   " "
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox Text6 
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
            Index           =   2
            Left            =   12105
            TabIndex        =   373
            Text            =   " "
            Top             =   240
            Width           =   930
         End
         Begin VB.CommandButton Command40 
            Caption         =   "..."
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
            Index           =   9
            Left            =   11145
            Picture         =   "AMASTOAHP.frx":5128
            TabIndex        =   372
            Top             =   240
            Width           =   175
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Visualizar costos"
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
            Index           =   3
            Left            =   11025
            TabIndex        =   371
            Top             =   915
            Width           =   2265
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   5100
            Index           =   4
            Left            =   60
            TabIndex        =   376
            Top             =   1320
            Visible         =   0   'False
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   8996
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   12648447
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   5100
            Index           =   3
            Left            =   60
            TabIndex        =   380
            Top             =   1320
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   8996
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   9465
            TabIndex        =   379
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   11265
            TabIndex        =   378
            Top             =   360
            Width           =   780
         End
         Begin VB.Label TEPRUEBA 
            AutoSize        =   -1  'True
            Caption         =   "TEPRUEBA"
            Height          =   300
            Left            =   345
            TabIndex        =   377
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame Frame15 
         Height          =   6495
         Index           =   5
         Left            =   -74880
         TabIndex        =   335
         Top             =   360
         Width           =   13335
         Begin VB.Frame Frame1 
            Caption         =   "A Ingresar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Index           =   2
            Left            =   5685
            TabIndex        =   359
            Top             =   120
            Width           =   2745
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " label303"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   218
               Left            =   1245
               TabIndex        =   365
               ToolTipText     =   "Doble Click Filtra Proyección Solo Compras"
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   208
               Left            =   300
               TabIndex        =   364
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label301"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   216
               Left            =   1245
               TabIndex        =   363
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Compras:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   206
               Left            =   300
               TabIndex        =   362
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label302"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   217
               Left            =   1245
               TabIndex        =   361
               Top             =   735
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fabricación:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   207
               Left            =   90
               TabIndex        =   360
               Top             =   840
               Width           =   1065
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Comprometido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Index           =   1
            Left            =   2850
            TabIndex        =   352
            Top             =   120
            Width           =   2745
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label203"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   215
               Left            =   1245
               TabIndex        =   358
               ToolTipText     =   "Doble Click Filtra Proyección Solo Pedidos"
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   205
               Left            =   300
               TabIndex        =   357
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Asignados y Reservas:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   204
               Left            =   90
               TabIndex        =   356
               Top             =   660
               Width           =   1065
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label202"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   214
               Left            =   1245
               TabIndex        =   355
               Top             =   735
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pedidos:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   203
               Left            =   300
               TabIndex        =   354
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label201"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   213
               Left            =   1245
               TabIndex        =   353
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Index           =   0
            Left            =   120
            TabIndex        =   345
            Top             =   120
            Width           =   2640
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label103"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   212
               Left            =   1260
               TabIndex        =   351
               Top             =   1155
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   202
               Left            =   315
               TabIndex        =   350
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Remoto:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   201
               Left            =   315
               TabIndex        =   349
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label102"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   211
               Left            =   1260
               TabIndex        =   348
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Propio:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   200
               Left            =   315
               TabIndex        =   347
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "label101"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   210
               Left            =   1260
               TabIndex        =   346
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Index           =   0
            Left            =   8520
            TabIndex        =   344
            Top             =   540
            Width           =   2325
         End
         Begin VB.Frame Frame15 
            Height          =   600
            Index           =   18
            Left            =   120
            TabIndex        =   339
            Top             =   5820
            Width           =   7455
            Begin VB.ComboBox CMBFILTRO 
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
               Left            =   4200
               Style           =   2  'Dropdown List
               TabIndex        =   341
               Top             =   160
               Width           =   3135
            End
            Begin VB.TextBox TXTBUSCAR 
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
               Left            =   720
               TabIndex        =   340
               Top             =   160
               Width           =   2535
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D8BD05&
               BackStyle       =   0  'Transparent
               Caption         =   "Campo:"
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
               Index           =   220
               Left            =   3360
               TabIndex        =   343
               Top             =   160
               Width           =   705
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D8BD05&
               BackStyle       =   0  'Transparent
               Caption         =   "Filtrar:"
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
               Index           =   209
               Left            =   120
               TabIndex        =   342
               Top             =   160
               Width           =   585
            End
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Visualizar costos"
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
            Index           =   2
            Left            =   10920
            TabIndex        =   338
            Top             =   1620
            Width           =   2265
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2445
            Index           =   0
            Left            =   120
            TabIndex        =   366
            Top             =   3360
            Width           =   13065
            _ExtentX        =   23045
            _ExtentY        =   4313
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   1245
            Index           =   1
            Left            =   120
            TabIndex        =   367
            Top             =   1920
            Width           =   13065
            _ExtentX        =   23045
            _ExtentY        =   2196
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   2445
            Index           =   2
            Left            =   120
            TabIndex        =   369
            Top             =   3360
            Visible         =   0   'False
            Width           =   13065
            _ExtentX        =   23045
            _ExtentY        =   4313
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   12648447
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Existencias"
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
            Left            =   8520
            TabIndex        =   368
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Stocks x Proveedores y Ubicaciones"
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
            Left            =   120
            TabIndex        =   337
            Top             =   1695
            Width           =   3615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Proyección"
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
            Index           =   3
            Left            =   120
            TabIndex        =   336
            Top             =   3135
            Width           =   3615
         End
      End
      Begin VB.Frame Frame15 
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
         Index           =   17
         Left            =   -74925
         TabIndex        =   138
         Top             =   6000
         Width           =   13260
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
            Left            =   12870
            TabIndex        =   175
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
            Left            =   12600
            Picture         =   "AMASTOAHP.frx":5432
            TabIndex        =   140
            Top             =   345
            Width           =   175
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"AMASTOAHP.frx":573C
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
            TabIndex        =   239
            Top             =   360
            Visible         =   0   'False
            Width           =   12390
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
            TabIndex        =   139
            Top             =   360
            Width           =   9630
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Height          =   6150
         Index           =   1
         Left            =   -74980
         TabIndex        =   122
         Top             =   600
         Width           =   13410
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
            Height          =   4035
            Left            =   70
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   187
            Top             =   720
            Width           =   10545
         End
         Begin VB.Frame Frame15 
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
            Index           =   16
            Left            =   10800
            TabIndex        =   176
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
               TabIndex        =   179
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
               TabIndex        =   178
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
               TabIndex        =   177
               Top             =   1365
               Width           =   1185
            End
         End
         Begin VB.Frame Frame15 
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
            Index           =   15
            Left            =   10800
            TabIndex        =   157
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
               TabIndex        =   162
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
               TabIndex        =   160
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
               TabIndex        =   163
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
               TabIndex        =   161
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
               TabIndex        =   159
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
               TabIndex        =   167
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
               TabIndex        =   166
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
               TabIndex        =   165
               Top             =   240
               Width           =   1155
            End
         End
         Begin VB.Frame Frame15 
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
            Index           =   14
            Left            =   10800
            TabIndex        =   156
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
               TabIndex        =   158
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   0
            Left            =   5700
            TabIndex        =   129
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   1
            Left            =   6855
            TabIndex        =   128
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   2
            Left            =   8025
            TabIndex        =   127
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Index           =   3
            Left            =   9315
            TabIndex        =   126
            Top             =   4770
            Width           =   1320
         End
         Begin VB.CommandButton Command46 
            Height          =   260
            Left            =   75
            TabIndex        =   125
            Top             =   4770
            Width           =   5580
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   70
            ScaleHeight     =   375
            ScaleWidth      =   10485
            TabIndex        =   123
            Top             =   315
            Width           =   10545
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
               TabIndex        =   124
               Top             =   105
               Width           =   10440
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
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   10680
         Begin VB.Frame Frame15 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Index           =   13
            Left            =   120
            TabIndex        =   328
            Top             =   3270
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
               TabIndex        =   329
               Top             =   315
               Width           =   4320
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Operaciones Standard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Index           =   12
            Left            =   120
            TabIndex        =   318
            Top             =   120
            Width           =   4530
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
               TabIndex        =   327
               Top             =   2520
               Width           =   845
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
               TabIndex        =   326
               Top             =   2520
               Width           =   845
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
               TabIndex        =   325
               Top             =   420
               Width           =   4320
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
               TabIndex        =   324
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   2520
               Width           =   550
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
               TabIndex        =   323
               ToolTipText     =   "Listado de TMU's"
               Top             =   180
               Visible         =   0   'False
               Width           =   690
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
               TabIndex        =   322
               ToolTipText     =   "Copia de TMU's"
               Top             =   180
               Visible         =   0   'False
               Width           =   690
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
               TabIndex        =   321
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
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
               TabIndex        =   320
               ToolTipText     =   "Elimina la Operación Activa"
               Top             =   2520
               Width           =   550
            End
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
               TabIndex        =   319
               ToolTipText     =   "Inserta Operación Alternativa a la Activa"
               Top             =   2520
               Width           =   550
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Detalle de la Operación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5925
            Index           =   7
            Left            =   4920
            TabIndex        =   240
            Top             =   120
            Width           =   5685
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
               TabIndex        =   297
               Top             =   4560
               Width           =   930
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
               TabIndex        =   296
               Top             =   4560
               Width           =   360
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
               TabIndex        =   295
               Top             =   4560
               Width           =   465
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
               TabIndex        =   294
               Top             =   3015
               Width           =   930
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
               TabIndex        =   293
               Top             =   3015
               Width           =   360
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
               TabIndex        =   292
               Top             =   3015
               Width           =   465
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
               TabIndex        =   291
               Top             =   2700
               Width           =   930
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
               TabIndex        =   290
               Top             =   2700
               Width           =   360
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
               TabIndex        =   289
               Top             =   2700
               Width           =   465
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
               TabIndex        =   288
               Top             =   2385
               Width           =   930
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
               TabIndex        =   287
               Top             =   2385
               Width           =   360
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
               TabIndex        =   286
               Top             =   2385
               Width           =   465
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
               TabIndex        =   285
               Top             =   1440
               Width           =   885
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
               TabIndex        =   284
               Top             =   405
               Width           =   1095
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
               TabIndex        =   283
               Top             =   720
               Width           =   4245
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
               TabIndex        =   282
               Top             =   1440
               Width           =   175
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
               TabIndex        =   281
               Text            =   " "
               Top             =   4560
               Width           =   675
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
               TabIndex        =   280
               Top             =   1440
               Width           =   1830
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
               TabIndex        =   279
               Top             =   3330
               Width           =   930
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
               TabIndex        =   278
               Top             =   3330
               Width           =   360
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
               TabIndex        =   277
               Top             =   3330
               Width           =   465
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
               TabIndex        =   276
               Text            =   " "
               Top             =   4560
               Width           =   675
            End
            Begin VB.Frame Frame15 
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
               Index           =   8
               Left            =   210
               TabIndex        =   273
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
                  TabIndex        =   274
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
                  TabIndex        =   275
                  ToolTipText     =   "Doble-Click para Ingresar Cadencia"
                  Top             =   280
                  Width           =   885
               End
            End
            Begin VB.Frame Frame15 
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
               Index           =   9
               Left            =   1800
               TabIndex        =   271
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
                  TabIndex        =   272
                  ToolTipText     =   "Porcentaje Demora Distribuida"
                  Top             =   285
                  Width           =   570
               End
            End
            Begin VB.Frame Frame15 
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
               Index           =   11
               Left            =   3290
               TabIndex        =   255
               Top             =   2280
               Width           =   2235
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
                  TabIndex        =   265
                  Top             =   360
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
                  TabIndex        =   264
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
                  Index           =   3
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   263
                  Top             =   980
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
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   262
                  Top             =   360
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
                  TabIndex        =   261
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
                  Index           =   3
                  Left            =   1935
                  TabIndex        =   260
                  Top             =   980
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
                  Index           =   4
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   259
                  Top             =   1290
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
                  Index           =   5
                  Left            =   1170
                  MaxLength       =   6
                  TabIndex        =   258
                  Top             =   1600
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
                  Index           =   4
                  Left            =   1935
                  TabIndex        =   257
                  Top             =   1290
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
                  Index           =   5
                  Left            =   1935
                  TabIndex        =   256
                  Top             =   1600
                  Width           =   175
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
                  TabIndex        =   270
                  Top             =   420
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
                  TabIndex        =   269
                  Top             =   730
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
                  TabIndex        =   268
                  Top             =   1040
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
                  TabIndex        =   267
                  Top             =   1330
                  Width           =   960
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
                  TabIndex        =   266
                  Top             =   1660
                  Width           =   960
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
               Picture         =   "AMASTOAHP.frx":57DE
               Style           =   1  'Graphical
               TabIndex        =   254
               ToolTipText     =   "Herramentales/Equipos Auxiliares"
               Top             =   4440
               Width           =   615
            End
            Begin VB.Frame Frame15 
               Caption         =   "RRHH - Personal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   950
               Index           =   10
               Left            =   0
               TabIndex        =   241
               Top             =   5000
               Width           =   5685
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
                  TabIndex        =   247
                  Top             =   550
                  Width           =   175
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
                  TabIndex        =   246
                  Top             =   550
                  Width           =   360
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
                  TabIndex        =   245
                  Top             =   550
                  Width           =   570
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
                  TabIndex        =   244
                  Top             =   240
                  Width           =   175
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
                  TabIndex        =   243
                  Top             =   240
                  Width           =   360
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
                  Index           =   0
                  Left            =   1215
                  TabIndex        =   242
                  Top             =   240
                  Width           =   570
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
                  TabIndex        =   253
                  Top             =   300
                  Width           =   330
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
                  TabIndex        =   252
                  Top             =   550
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
                  TabIndex        =   251
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
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   250
                  Top             =   240
                  Width           =   2460
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
                  TabIndex        =   249
                  Top             =   300
                  Width           =   855
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
                  TabIndex        =   248
                  Top             =   600
                  Width           =   330
               End
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
               TabIndex        =   317
               ToolTipText     =   "Doble Click para Detalle Mant.Herramental"
               Top             =   4560
               Width           =   1125
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
               TabIndex        =   316
               Top             =   1035
               Width           =   2460
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
               TabIndex        =   315
               Top             =   1130
               Width           =   810
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
               TabIndex        =   314
               Top             =   1035
               Width           =   885
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
               TabIndex        =   313
               Top             =   405
               Width           =   570
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
               TabIndex        =   312
               Top             =   2160
               Width           =   1800
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
               TabIndex        =   311
               ToolTipText     =   "Doble-Click para Abrir Tiempos Variables por Lote"
               Top             =   3110
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
               TabIndex        =   310
               Top             =   2795
               Width           =   1095
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
               TabIndex        =   309
               ToolTipText     =   "Doble-Click para Abrir Tiempos de Preparación"
               Top             =   2480
               Width           =   1095
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
               TabIndex        =   308
               Top             =   1564
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
               TabIndex        =   307
               Top             =   1755
               Width           =   4245
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
               TabIndex        =   306
               Top             =   500
               Width           =   960
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   420
               Picture         =   "AMASTOAHP.frx":5AE8
               ToolTipText     =   "Abrir Tiempos de Preparación"
               Top             =   2070
               Width           =   480
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
               TabIndex        =   305
               ToolTipText     =   "Tiempos por Unidad"
               Top             =   405
               Width           =   165
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
               TabIndex        =   304
               Top             =   3425
               Width           =   1095
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
               TabIndex        =   303
               Top             =   11785
               Width           =   750
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
               TabIndex        =   302
               Top             =   1545
               Width           =   750
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
               TabIndex        =   301
               Top             =   405
               Width           =   300
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
               TabIndex        =   300
               ToolTipText     =   "Doble Click para Detalle Mant.Herramental"
               Top             =   4590
               Width           =   945
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
               TabIndex        =   299
               Top             =   1110
               Width           =   495
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
               TabIndex        =   298
               Top             =   480
               Width           =   1275
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
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6160
         Index           =   2
         Left            =   -74980
         TabIndex        =   41
         Top             =   600
         Width           =   13440
         Begin VB.PictureBox Picture2 
            Height          =   5535
            Left            =   10200
            ScaleHeight     =   5475
            ScaleWidth      =   3195
            TabIndex        =   330
            Top             =   360
            Width           =   3255
            Begin VB.Image Image3 
               Height          =   5535
               Left            =   0
               Picture         =   "AMASTOAHP.frx":5DF2
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3255
            End
         End
         Begin VB.Frame Frame15 
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
            Index           =   4
            Left            =   360
            TabIndex        =   188
            Top             =   3800
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
               TabIndex        =   215
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
                  TabIndex        =   218
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
                  TabIndex        =   217
                  Top             =   475
                  Width           =   175
               End
               Begin VB.CheckBox Check2 
                  Caption         =   "x Millar"
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
                  Height          =   225
                  Left            =   210
                  TabIndex        =   216
                  Top             =   1785
                  Width           =   960
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "U-M: "
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
                  Index           =   52
                  Left            =   3570
                  TabIndex        =   238
                  Top             =   525
                  Width           =   480
               End
               Begin VB.Label Label10 
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
                  Index           =   54
                  Left            =   6510
                  TabIndex        =   237
                  Top             =   525
                  Width           =   270
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Uso:                       D:"
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
                  Index           =   53
                  Left            =   4830
                  TabIndex        =   236
                  Top             =   525
                  Width           =   375
               End
               Begin VB.Label Label23 
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
                  Height          =   540
                  Left            =   1260
                  TabIndex        =   235
                  Top             =   1455
                  Width           =   5895
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
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
                  Height          =   270
                  Left            =   6825
                  TabIndex        =   234
                  Top             =   480
                  Width           =   330
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
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
                  Height          =   270
                  Left            =   5250
                  TabIndex        =   233
                  Top             =   480
                  Width           =   1170
               End
               Begin VB.Label Label19 
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
                  Height          =   270
                  Left            =   1260
                  TabIndex        =   232
                  Top             =   480
                  Width           =   2050
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código:                                          U-M:           Uso:                       D:"
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
                  Left            =   540
                  TabIndex        =   231
                  Top             =   525
                  Width           =   690
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
                  Index           =   17
                  Left            =   105
                  TabIndex        =   230
                  Top             =   840
                  Width           =   1065
               End
               Begin VB.Label Label9 
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
                  Height          =   270
                  Left            =   4200
                  TabIndex        =   229
                  Top             =   480
                  Width           =   435
               End
               Begin VB.Label Label4 
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
                  Left            =   1260
                  TabIndex        =   228
                  Top             =   840
                  Width           =   5895
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Referencia:"
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
                  Index           =   16
                  Left            =   105
                  TabIndex        =   227
                  Top             =   1450
                  Width           =   1065
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
                  TabIndex        =   226
                  Top             =   885
                  Width           =   1065
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
                  TabIndex        =   225
                  Top             =   885
                  Width           =   570
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
                  TabIndex        =   224
                  Top             =   840
                  Width           =   945
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
                  TabIndex        =   223
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
                  TabIndex        =   222
                  Top             =   840
                  Width           =   750
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
                  TabIndex        =   221
                  Top             =   885
                  Width           =   645
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
                  TabIndex        =   220
                  Top             =   840
                  Width           =   855
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
                  Left            =   1260
                  TabIndex        =   219
                  Top             =   840
                  Width           =   1095
               End
            End
            Begin VB.TextBox Text39 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   1260
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   200
               Top             =   1450
               Width           =   5910
            End
            Begin VB.CommandButton Command27 
               Caption         =   "App"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7400
               TabIndex        =   199
               ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
               Top             =   470
               Width           =   550
            End
            Begin VB.TextBox Text32 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5250
               TabIndex        =   198
               Text            =   " "
               Top             =   475
               Width           =   1170
            End
            Begin VB.TextBox Text31 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1260
               TabIndex        =   197
               Top             =   475
               Width           =   2115
            End
            Begin VB.CommandButton Command26 
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
               Height          =   375
               Left            =   7400
               TabIndex        =   196
               ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
               Top             =   855
               Width           =   550
            End
            Begin VB.TextBox Text27 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   6825
               TabIndex        =   195
               Text            =   " "
               Top             =   475
               Width           =   330
            End
            Begin VB.CommandButton Command19 
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
               Height          =   375
               Left            =   7400
               TabIndex        =   194
               ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
               Top             =   1260
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
               TabIndex        =   193
               Top             =   475
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
               Height          =   375
               Left            =   7400
               TabIndex        =   192
               ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
               Top             =   1650
               Width           =   550
            End
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
               Left            =   210
               TabIndex        =   191
               Top             =   1785
               Width           =   960
            End
            Begin VB.TextBox Text41 
               Alignment       =   1  'Right Justify
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
               Left            =   1260
               TabIndex        =   190
               Text            =   " "
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox Text42 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   189
               Text            =   " "
               Top             =   840
               Width           =   750
            End
            Begin VB.Label Label10 
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
               Index           =   57
               Left            =   6510
               TabIndex        =   214
               Top             =   525
               Width           =   270
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Uso:                       D:"
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
               Index           =   56
               Left            =   4830
               TabIndex        =   213
               Top             =   525
               Width           =   375
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "U-M:"
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
               Index           =   55
               Left            =   3570
               TabIndex        =   212
               Top             =   525
               Width           =   480
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Referencia:"
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
               TabIndex        =   211
               Top             =   1450
               Width           =   1065
            End
            Begin VB.Label Label40 
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
               Height          =   540
               Left            =   1260
               TabIndex        =   210
               Top             =   840
               Width           =   5895
            End
            Begin VB.Label Label41 
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
               Height          =   270
               Left            =   4200
               TabIndex        =   209
               Top             =   480
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
               TabIndex        =   208
               Top             =   840
               Width           =   1065
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
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
               Left            =   540
               TabIndex        =   207
               Top             =   525
               Width           =   690
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
               Index           =   27
               Left            =   120
               TabIndex        =   206
               Top             =   885
               Width           =   1065
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
               Index           =   24
               Left            =   5625
               TabIndex        =   205
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6210
               TabIndex        =   204
               Top             =   840
               Width           =   945
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
               Index           =   25
               Left            =   4050
               TabIndex        =   203
               Top             =   885
               Width           =   645
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
               Index           =   26
               Left            =   2400
               TabIndex        =   202
               Top             =   885
               Width           =   645
            End
            Begin VB.Label Label42 
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
               TabIndex        =   201
               Top             =   840
               Width           =   855
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
            Left            =   9585
            TabIndex        =   134
            ToolTipText     =   "Ordenamiento de Estructura"
            Top             =   405
            Width           =   525
         End
         Begin VB.Frame Frame15 
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
            Index           =   3
            Left            =   8565
            TabIndex        =   69
            Top             =   4740
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   181
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
         Begin VB.Frame Frame15 
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
            Index           =   2
            Left            =   8565
            TabIndex        =   46
            Top             =   3195
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
            Left            =   8565
            Picture         =   "AMASTOAHP.frx":1B5E34
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   405
            Width           =   1005
         End
         Begin VB.Frame Frame15 
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
            Index           =   1
            Left            =   8565
            TabIndex        =   44
            Top             =   2025
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
               TabIndex        =   68
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
               TabIndex        =   67
               ToolTipText     =   "Nueva Fórmula por Copia de una Pre-Existente"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame Frame15 
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
            Index           =   0
            Left            =   8565
            TabIndex        =   42
            Top             =   855
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
               TabIndex        =   135
               Top             =   720
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
         Height          =   6255
         Index           =   0
         Left            =   20
         TabIndex        =   11
         Top             =   480
         Width           =   13545
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
            Height          =   135
            Index           =   2
            Left            =   6840
            TabIndex        =   21
            Top             =   6240
            Visible         =   0   'False
            Width           =   6600
            Begin VB.TextBox Text28 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   3720
               TabIndex        =   73
               Top             =   1005
               Width           =   1150
            End
            Begin VB.TextBox Text21 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   3465
               TabIndex        =   39
               Top             =   1890
               Width           =   1410
            End
            Begin VB.TextBox Text13 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1260
               TabIndex        =   25
               Top             =   525
               Width           =   1410
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   183
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
               TabIndex        =   154
               Top             =   2785
               Width           =   660
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   3510
               TabIndex        =   153
               Top             =   2700
               Width           =   1410
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1260
               TabIndex        =   152
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
               TabIndex        =   151
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
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1260
               TabIndex        =   34
               Top             =   1320
               Width           =   3615
            End
            Begin VB.Label Label33 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
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
            Left            =   6840
            TabIndex        =   100
            Top             =   285
            Width           =   6600
            Begin VB.TextBox Text52 
               Alignment       =   1  'Right Justify
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
               Left            =   2400
               MaxLength       =   4
               TabIndex        =   172
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
               Left            =   1575
               TabIndex        =   171
               ToolTipText     =   "Actualizacion Masiva por Proveedor"
               Top             =   2010
               Width           =   180
            End
            Begin VB.TextBox Text51 
               Alignment       =   1  'Right Justify
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
               Left            =   1200
               MaxLength       =   2
               TabIndex        =   170
               ToolTipText     =   "Cantidad dias Lead Time"
               Top             =   2010
               Width           =   375
            End
            Begin VB.TextBox Text49 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   4680
               TabIndex        =   145
               Top             =   1350
               Width           =   1305
            End
            Begin VB.CheckBox Check3 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   430
               Left            =   240
               TabIndex        =   130
               Top             =   1065
               Width           =   1905
            End
            Begin VB.TextBox Text40 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   4680
               TabIndex        =   120
               Top             =   950
               Width           =   1305
            End
            Begin VB.TextBox Text18 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   4680
               TabIndex        =   115
               Top             =   240
               Width           =   1305
            End
            Begin VB.TextBox Text19 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   4680
               TabIndex        =   109
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
               TabIndex        =   103
               Top             =   300
               Width           =   175
            End
            Begin VB.TextBox Text20 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   4680
               TabIndex        =   102
               Top             =   1710
               Width           =   1305
            End
            Begin VB.TextBox Text8 
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
               Height          =   285
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   101
               Text            =   " "
               Top             =   300
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
               Left            =   1860
               TabIndex        =   173
               ToolTipText     =   "Autonomía Proyectada de Cálculo"
               Top             =   2130
               Width           =   495
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
               Left            =   705
               TabIndex        =   169
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   2130
               Width           =   495
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
               Height          =   165
               Left            =   525
               TabIndex        =   164
               Top             =   1860
               Width           =   600
            End
            Begin VB.Label Label71 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   240
               TabIndex        =   155
               Top             =   1860
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
               Left            =   4155
               TabIndex        =   150
               Top             =   2145
               Width           =   540
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   4680
               TabIndex        =   149
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
               Left            =   3240
               TabIndex        =   146
               Top             =   1455
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
               Height          =   195
               Left            =   525
               TabIndex        =   133
               Top             =   1530
               Width           =   1170
            End
            Begin VB.Label Label58 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   240
               TabIndex        =   132
               Top             =   1525
               Width           =   195
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
               Left            =   3495
               TabIndex        =   121
               Top             =   1065
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
               Left            =   3495
               TabIndex        =   108
               Top             =   1785
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
               Left            =   3495
               TabIndex        =   107
               Top             =   345
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
               Left            =   3495
               TabIndex        =   106
               Top             =   690
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
               TabIndex        =   105
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label Label32 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   240
               TabIndex        =   104
               Top             =   720
               Width           =   1920
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Características"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5820
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   285
            Width           =   6600
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
               Left            =   6075
               MaxLength       =   1
               TabIndex        =   147
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   1335
               Width           =   375
            End
            Begin VB.TextBox Text48 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   142
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
               TabIndex        =   141
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
               Left            =   5400
               TabIndex        =   137
               ToolTipText     =   "Actualización y Listado"
               Top             =   2850
               Width           =   1020
            End
            Begin VB.TextBox Text44 
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
               Left            =   1155
               TabIndex        =   112
               Top             =   2850
               Width           =   4170
            End
            Begin VB.TextBox Text46 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   118
               Top             =   4455
               Width           =   2175
            End
            Begin VB.TextBox Text16 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   116
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
               TabIndex        =   113
               Top             =   1820
               Width           =   175
            End
            Begin VB.TextBox Text35 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   111
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
               TabIndex        =   110
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
               Height          =   360
               Left            =   6135
               TabIndex        =   97
               Top             =   5040
               Width           =   175
            End
            Begin VB.TextBox Text25 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   4320
               TabIndex        =   96
               Text            =   " "
               Top             =   5040
               Width           =   1815
            End
            Begin VB.TextBox Text36 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   95
               Top             =   5040
               Width           =   1620
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
               Height          =   360
               Left            =   6210
               TabIndex        =   94
               Top             =   3360
               Width           =   180
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
               Height          =   360
               Left            =   6210
               TabIndex        =   93
               Top             =   3900
               Width           =   180
            End
            Begin VB.TextBox Text23 
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
               Left            =   1155
               TabIndex        =   64
               Top             =   3360
               Width           =   5040
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
               Left            =   1155
               TabIndex        =   63
               Top             =   3900
               Width           =   5040
            End
            Begin VB.TextBox Text24 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   4890
               TabIndex        =   20
               Top             =   4455
               Width           =   1425
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
               TabIndex        =   16
               Top             =   1320
               Width           =   175
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   15
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
               TabIndex        =   14
               Top             =   380
               Width           =   175
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Index           =   0
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   13
               Text            =   " "
               Top             =   360
               Width           =   570
            End
            Begin VB.Label Label57 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1995
               TabIndex        =   168
               Top             =   2325
               Width           =   4455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dec:"
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
               Index           =   102
               Left            =   5430
               TabIndex        =   148
               ToolTipText     =   "Cantidad de Decimales"
               Top             =   1425
               Width           =   615
            End
            Begin VB.Label Label64 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1995
               TabIndex        =   144
               Top             =   840
               Width           =   4455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo de Familias:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   29
               Left            =   120
               TabIndex        =   143
               Top             =   765
               Width           =   900
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cod.Intern.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   0
               TabIndex        =   136
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   19
               Left            =   195
               TabIndex        =   119
               Top             =   4400
               Width           =   870
            End
            Begin VB.Label Label31 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1995
               TabIndex        =   114
               Top             =   1815
               Width           =   4455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Revisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   49
               Left            =   3405
               TabIndex        =   99
               Top             =   4965
               Width           =   840
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   48
               Left            =   -1575
               TabIndex        =   98
               Top             =   5170
               Width           =   2655
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Contable:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   100
               Left            =   150
               TabIndex        =   92
               Top             =   1720
               Width           =   930
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo I.V.A.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   101
               Left            =   120
               TabIndex        =   91
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   3
               Left            =   105
               TabIndex        =   66
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   2
               Left            =   105
               TabIndex        =   65
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
                  Weight          =   400
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
                  Weight          =   400
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   4
               Left            =   2550
               TabIndex        =   36
               Top             =   4395
               Width           =   2070
            End
            Begin VB.Label Label30 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1995
               TabIndex        =   18
               Top             =   1335
               Width           =   3540
            End
            Begin VB.Label Label29 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1995
               TabIndex        =   17
               Top             =   375
               Width           =   4455
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   3250
            Index           =   10
            Left            =   6840
            TabIndex        =   334
            Top             =   2880
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   5741
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   100
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "No Usados"
         Height          =   4215
         Left            =   525
         TabIndex        =   53
         Top             =   1230
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            Index           =   0
            Left            =   3465
            TabIndex        =   62
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
            TabIndex        =   61
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
            Index           =   0
            Left            =   210
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   55
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
            TabIndex        =   54
            Top             =   210
            Width           =   1290
         End
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   -67680
         TabIndex        =   333
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Left            =   -68280
         TabIndex        =   331
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Index           =   0
         Left            =   -71205
         TabIndex        =   9
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   8
         Top             =   3960
         Width           =   15
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "AMASTOAHP.frx":1B613E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "AMASTOAHP.frx":1B6372
      TabIndex        =   186
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
      TabIndex        =   90
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
      Height          =   270
      Index           =   40
      Left            =   6675
      TabIndex        =   88
      Top             =   600
      Width           =   525
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   105
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6900
      TabIndex        =   35
      Top             =   810
      Width           =   585
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   315
      TabIndex        =   10
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      Begin VB.Menu mnuInic_novedades 
         Caption         =   "Inicializar Novedades"
         Visible         =   0   'False
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
         Visible         =   0   'False
      End
      Begin VB.Menu phl4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SustCod 
         Caption         =   "Sustitución de Código de Articulo"
      End
      Begin VB.Menu BaLog 
         Caption         =   "Baja Lógica Individual"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ubicacionxprov 
         Caption         =   "Ubicaciones por Proveedor"
         Begin VB.Menu stockdetallexproveedor 
            Caption         =   "Stock Detalle x Proveedor"
         End
         Begin VB.Menu soloxproveedor 
            Caption         =   "Por proveedor"
         End
      End
   End
   Begin VB.Menu ListCon 
      Caption         =   "Listados y Consultas"
      Begin VB.Menu mnuinformeCostosVsVentas 
         Caption         =   "Informe Costos vs Ventas"
      End
      Begin VB.Menu GralMa 
         Caption         =   "Generales del Maestro"
         Begin VB.Menu AtGral 
            Caption         =   "Atributos Generales"
         End
         Begin VB.Menu StoLot 
            Caption         =   "Stock Minino y Lote de Reposición"
            Visible         =   0   'False
         End
         Begin VB.Menu CapInv 
            Caption         =   "Captura para Inventarios"
         End
      End
      Begin VB.Menu CosRep 
         Caption         =   "Costos de Reposición"
         Begin VB.Menu NumPro 
            Caption         =   "Por Numero de Proveedor"
            Visible         =   0   'False
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
      Begin VB.Menu mnukardex 
         Caption         =   "Ficha Kardex"
      End
      Begin VB.Menu LisPre 
         Caption         =   "Lista de Precios de Venta"
         Visible         =   0   'False
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
Attribute VB_Name = "AMASTOAHP"
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

Sub ACTUFORMUPAT(Ci%, LOTEANTERIOR, LOTENUEVO)
     '
     CONDI$ = "CodIConj = " & Ci%
     Call ACTUREGISTRO("FORMUPAT", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
     
     'VALIDO SI EXISTE EN LA FORMULA PATRON, SI NO SE VA.
     If RAFE& < 1 Then Exit Sub
     '
     SQLX$ = "SELECT  * FROM FORMULAS WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CodIConj = " & Ci%
     Dim FORMUTMP As New ADODB.Recordset
     Set FORMUTMP = FLEXCONN.Execute(SQLX$)
     '
     With FORMUTMP
        Do While Not .EOF
           Uso# = XVALO(!usobruto)
           Uso# = Uso# * LOTENUEVO
           CIXXII% = XVALO(!codielem)
           CODELEM$ = SAFETEXT$(!CodXElem)
           NUOR% = XVALO(!NUORIT)
           
           'AGREGO A LA FORMULA PATRON EL NUEVO USO
10         RNDX$ = RANDSTRING$(8)
           If CantRegistros("FORMUPAT", "CODXCONJ = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
           '
           Call CreaRegistro("FORMUPAT", "CODXCONJ", RNDX$, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODICONJ", "CODXCONJ = '" & RNDX$ & "'", Ci%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODIELEM", "CODXCONJ = '" & RNDX$ & "'", CIXXII%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODXELEM", "CODXCONJ = '" & RNDX$ & "'", CODEXT$(CIXXII%), RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "USOBRUTO", "CODXCONJ = '" & RNDX$ & "'", FORMATNUM$(Uso#, "F16.5"), RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "NUORIT", "CODXCONJ = '" & RNDX$ & "'", NUOR%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "CODIEMPR", "CODXCONJ = '" & RNDX$ & "'", CodiEmp%, RAFE&, "NOMESS")
           Call ACTUREGISTRO("FORMUPAT", "MARBORRA", "CODXCONJ = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
           '
           Call ACTUREGISTRO("FORMUPAT", "CODXCONJ", "CODXCONJ = '" & RNDX$ & "'", CODEXT$(Ci%), RAFE&, "NOMESS")
           '
        .MoveNext
        Loop
     End With
     FORMUTMP.Close
     Set FORMUTMP = Nothing
     '
     'BORRO LOS REGISTROS MARCADOS CON 1
     CONDI$ = "CodIConj = " & Ci% & " AND MARBORRA = 1 "
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
      CONDI$ = "CODINT = " & Ci%
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

Sub CARGACOSTOXPROVE()
    MSF(10).Rows = 1: J& = 0
    CI1% = XVALO(Label28)
    
    If CI1% < 1 Then Exit Sub
    SQLX$ = "SELECT o.Cod_Inte,o.Npro_Ocom, o.RaSo_Prov, oc.Via_Entrega FROM OCOMDETA o WITH(NOLOCK) INNER JOIN OCOMENCA OC"
    SQLX$ = SQLX$ + " ON O.nume_ocom = OC.nume_ocom WHERE O.Cod_Inte = " & CI1% & " and O.stat_ocom < 9 and O.stat_ocom >= 0"
    SQLX$ = SQLX$ + " and  Cant_Rec > 0 "
    SQLX$ = SQLX$ + " group by O.COD_INTE, O.NPRO_OCOM,O.RaSo_Prov,OC.VIA_ENTREGA "
    SQLX$ = SQLX$ + " ORDER BY O.Npro_Ocom,OC.Via_Entrega DESC"
    '
    '
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        CTAPROV = XVALO(!Npro_Ocom)
        RACLI$ = SAFETEXT$(!Raso_Prov)
        VIAENTR% = XVALO(!Via_Entrega)
        J& = J& + 1
        MSF(10).Rows = J& + 1
        MSF(10).TextMatrix(J&, 1) = CTAPROV
        MSF(10).TextMatrix(J&, 2) = RACLI$
        MSF(10).TextMatrix(J&, 3) = ECOARCH$("VIASENTR", MKI$(VIAENTR%))
        MSF(10).TextMatrix(J&, 8) = VIAENTR%
        .MoveNext
      Loop
    End With
    '
    NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))

    For i& = 1 To MSF(10).Rows - 1
       CTAPROV = XVALO(MSF(10).TextMatrix(i&, 1))
       'ESTAS 2 LINEAS SIGUIENTES SON PARA OBTENER EL LOTE
       VIAENTR% = XVALO(MSF(10).TextMatrix(i&, 8))
       NUOC& = ULTIMACOMPRA_SEGUN_PROV_VIA&(CI1%, CTAPROV, VIAENTR%)
       'FIN LOTE
       If NUOC& < 1 Then GoTo 80
       
       LOTE$ = LOTESEGUNCODIGO_PROV_COMPRA$(CI1%, CTAPROV, NUOC&)
       
       'A PARTIR DE AQUI MODIFICADO MINEPART
       InfoProveMineparts% = 0
       If NprovMineparts = CTAPROV Then
          'si el proveedor es minepart trae los datos de via,ctapro,pjenancionalizacion del exterior
          InfoProveMineparts% = CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, ProvExterior, ViaEntExt, PjeNacion, CoefNacion)
       End If
       'FIN MODIFICADO MINEPART
       
        COSFOB# = ROUND(COSTOFOB#(LOTE$), 2) ' TRAE EL COSTO DE LA ORDEN DE COMPRA
        MONECOMP% = MONELOTE%(LOTE$)
        COEFCON = 1
        If MONECOMP% < 2 Then
         'ESTO ES PARA QEU CUANDO SE COMPRA EN $ TENGA EL TIPO DE CAM DE LA OC Y LO PASE A U$S
         TipoCambioCom = TipoCambioCompra(LOTE$)
         If TipoCambioCom > 0 Then
          COEFCON = TICAMONE(MONECOMP%) / TipoCambioCom
         Else
          'COEFCON = TICAMONE(MONECOMP%) / TICAMONE(MONEMI%)
          COEFCON = TICAMONE(MONECOMP%) / 9.42 'ultimo valor antes de antes de guardar el valor de la cotizadcion en la o.c
         End If
        End If
        COSFOB# = COEFCON * COSFOB#
        NUMLOTE& = XVALO(Mid$(LOTE$, 4))
        If InfoProveMineparts% < 1 Then
          PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NUMLOTE&, , MONECOMP%)
          GASNAC# = XVALO(COSFOB# * PORCEN_GASNAC#)
        Else
          'CASO PROVEEDOR MINEPARTS
          COEF_GASNAC# = CoefNacion 'VALOR DEVUELTO DE FUNCION InfoProveMineparts%
          GASNAC# = XVALO(COSFOB# * COEF_GASNAC#)
          MSF(10).TextMatrix(i&, 1) = ProvExterior
          MSF(10).TextMatrix(i&, 2) = rasocli$((-1) * ProvExterior)
          CTAPROV = ProvExterior 'OJO ACA ACTUALIZA EL NRO DE PROVEEDOR SI ERA MINEPARTS
        End If
        
        MODONACIONALIZACION% = 0
        MSF(10).TextMatrix(i&, 4) = FORMATNUM$(COSFOB#, "F9.2")
        MSF(10).TextMatrix(i&, 5) = FORMATNUM$(GASNAC#, "F9.2")
        If GASNAC# <= 0 Then
          GASNAC# = COSFOB#
          MODONACIONALIZACION% = 1
        End If
           
        PREVENTAX# = PRECIOVENTA_X_FORMULA(XVALO(GASNAC#), VIAENTR%, CI1%, MODONACIONALIZACION%, CTAPROV, (-1))
        MSF(10).TextMatrix(i&, 6) = FORMATNUM$(PREVENTAX#, "F12.2")
        VENTAEDITADO# = PRECIODEVENTA_AHP(CI1%, CTAPROV, VIAENTR%)
        If VENTAEDITADO# > 0 Then
          MSF(10).TextMatrix(i&, 7) = FORMATNUM$(VENTAEDITADO#, "F12.2")
          If ROUND(VENTAEDITADO#, 2) < ROUND(MSF(10).TextMatrix(i&, 6), 2) Then
            MSF(10).COL = 7
            MSF(10).Row = i&
            MSF(10).CellForeColor = vbRed
          End If
        Else
           MSF(10).TextMatrix(i&, 7) = MSF(10).TextMatrix(i&, 6)
        End If
80  Next i&
    
 
End Sub
Sub CARGAR_COSTOS_AHP()

End Sub

Sub CARGAR_STOCKS()
   CIX1% = XVALO(Label15(2))
   Call MUEDISAHP(CIX1%)
End Sub



Sub COSTOFOBYNACIONALIZADO(LOTE$, RET_FOB#, RET_GASNAC#)
        RET_FOB# = 0
        RET_GASNAC# = 0
        CTAPROV = PROVELOTE%(LOTE$)
        NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
      'A PARTIR DE AQUI MODIFICADO MINEPART
       If NprovMineparts = CTAPROV Then
          'si el proveedor es minepart trae los datos de via,ctapro,pjenancionalizacion del exterior
          InfoProveMineparts% = CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, ProvExterior, ViaEntExt, PjeNacion, CoefNacion)
       End If
       'FIN MODIFICADO MINEPART
       
        COSFOB# = ROUND(COSTOFOB#(LOTE$), 2) ' TRAE EL COSTO DE LA ORDEN DE COMPRA
        MONECOMP% = MONELOTE%(LOTE$)
        COEFCON = 1
        If MONECOMP% < 2 Then
         'ESTO ES PARA QEU CUANDO SE COMPRA EN $ TENGA EL TIPO DE CAM DE LA OC Y LO PASE A U$S
         TipoCambioCom = TipoCambioCompra(LOTE$)
         If TipoCambioCom > 0 Then
          COEFCON = TICAMONE(MONECOMP%) / TipoCambioCom
         Else
          'COEFCON = TICAMONE(MONECOMP%) / TICAMONE(MONEMI%)
          COEFCON = TICAMONE(MONECOMP%) / 9.42 'ultimo valor antes de antes de guardar el valor de la cotizadcion en la o.c
         End If
        End If
        COSFOB# = COEFCON * COSFOB#
        NUMLOTE& = XVALO(Mid$(LOTE$, 4))
        If InfoProveMineparts% < 1 Then
          PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NUMLOTE&, , MONECOMP%)
          GASNAC# = XVALO(COSFOB# * PORCEN_GASNAC#)
        Else
          'CASO PROVEEDOR MINEPARTS
          COEF_GASNAC# = CoefNacion 'VALOR DEVUELTO DE FUNCION InfoProveMineparts%
          GASNAC# = XVALO(COSFOB# * COEF_GASNAC#)
          MSF(10).TextMatrix(i&, 1) = ProvExterior
          MSF(10).TextMatrix(i&, 2) = rasocli$((-1) * ProvExterior)
          CTAPROV = ProvExterior 'OJO ACA ACTUALIZA EL NRO DE PROVEEDOR SI ERA MINEPARTS
        End If
        
        MODONACIONALIZACION% = 0
'        MSF(10).TextMatrix(i&, 4) = FORMATNUM$(COSFOB#, "F9.2")
'        MSF(10).TextMatrix(i&, 5) = FORMATNUM$(GASNAC#, "F9.2")
        If GASNAC# <= 0 Then
          GASNAC# = COSFOB#
          MODONACIONALIZACION% = 1
        End If
        RET_FOB# = COSFOB#
        RET_GASNAC# = XVALO(GASNAC#)
        
End Sub

Sub FobPorcgastoPreventaProveedVia(LOTE$, CI1%, DevFob#, DevFobNacionalizado#, DevPreventa#, DevProveedor, DevViaentrega%, DevMONECOMP%)
             If TRIM$(LOTE$) = "" Then Exit Sub
             '
             VIALOTEXX% = VIALOTE%(LOTE$)
             
             VIA_ENTREGAX$ = TRIM$(ECOARCH$("VIASENTR", MKI$(VIALOTEXX%)))
             
             'COSTO FOB
             COSFOB# = COSTOFOB#(LOTE$)
             MONECOMP% = MONELOTE%(LOTE$)
             COEFCON = 1
             If MONECOMP% <> 2 Then
                'ESTO ES PARA QEU CUANDO SE COMPRA EN $ TENGA EL TIPO DE CAM DE LA OC Y LO PASE A U$S
                TipoCambioCom = TipoCambioCompra(LOTE$)
                If TipoCambioCom > 0 Then
                   COEFCON = TICAMONE(MONECOMP%) / TipoCambioCom
                Else
'                   COEFCON = TICAMONE(MONECOMP%) / TICAMONE(MONEMI%)
                   COEFCON = TICAMONE(MONECOMP%) / 9.42 'ultimo valor antes de antes de guardar el valor de la cotizadcion en la o.c
                End If
             End If
             COSFOB# = COEFCON * COSFOB#
             KOSTO$ = FORMATNUM$(COSFOB#, "F12.2")
             '
             NUMLOTE& = XVALO(Mid$(LOTE$, 4))
             PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NUMLOTE&, , MONECOMP%)
             GASNAC# = XVALO(COSFOB# * PORCEN_GASNAC#)
             MODONACIONALIZACION% = 0

             If GASNAC# <= 0 Then
               GASNAC# = COSFOB#
               MODONACIONALIZACION% = 1
             End If

             PREVENTAX# = PRECIOVENTA_X_FORMULA(XVALO(GASNAC#), VIALOTEXX%, CI1%, MODONACIONALIZACION%, NPROV)
             '
             '
             'PARAMETROS QUE DEVUELVE LA RUTINA
             DevFob = COSFOB#
             DevFobNacionalizado# = GASNAC#
             DevPreventa# = PREVENTAX#
             DevProveedor = PROVELOTE%(LOTE$)
             DevViaentrega% = VIALOTEXX%
             DevMONECOMP% = MONECOMP%
End Sub

Sub MUEDISAHP(CIXI%)
     '
     Static CTX%, VECOPEN$, CMPED$, HOYY, MODODOSI%
     '
     CIXI% = CODINT%(Text1)
     CADECI% = XVALO(CNTDC$(CIXI%))

     Screen.MousePointer = 11
     If CTX% <> 1 Then
'        IDENSEDE$ = CONTROL$("", "IDENSEDE")
'        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(Control$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        HOYY = HOY(HOS$)
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "DOSIVA") > 0 Then MODODOSI% = 1
        CTX% = 1
     End If
     '
     MSF(0).Rows = 1
'     MSF(1).Rows = 1
     List4(0).Clear
     STOPROP = 0
     STOREMO = 0
     TOPEDID = 0
     TOSUMIN = 0
     TOFABRI = 0
     TORESER = 0
     STOMINI = STOMIN(CIXI%)
     LECONOM = LOTESTAN(CIXI%)
     GoSub PRESEDIS
     '
     Call CARDEPOS
     '
     Screen.MousePointer = 11
     Call VESARRAS(CIXI%)
     '
     Screen.MousePointer = 11
     STOPROP = 0
     If CIXI% > 0 Then
       For U& = 1 To NUSU%
         CAPAR = STODEPOS(CIXI%, COSU%(U&)) '        XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
         If Abs(CAPAR) >= 0.05 Then
           XXX$ = ADEPO$(U&) + CSTRING$(MKS$(CAPAR), 4, 11, CADECI%, 2)
           If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
           XXX$ = AJUSTI$(XXX$, 60) + FORMATNUM$(COSU%(U&), "F4.0")
           List4(0).AddItem XXX$
           If DISTO%(U&) >= 0 Then
             STOPROP = STOPROP + CAPAR
           End If
         End If
       Next U&
       GoSub PRESEDIS
     End If
     '
10   MSF(0).Rows = 1
     List4(1).Clear
     HOII% = HOY(HO$)
     LI2TEX$ = Space$(156)
     Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 29)
     Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, CADECI%, 2), 44, 9)
     Call REPLA(LI2TEX$, HO$, 74, 9)
     
     OBSER$ = ""
     If STOPROP < 0 Then
          OBSER$ = "!!! Negativo !!!"
        ElseIf STOPROP < STOMINI Then
          OBSER$ = "Sub-Mínimo"
     End If
     Call REPLA(LI2TEX$, OBSER$, 84, 24)
     List4(1).AddItem LI2TEX$
     SALSTO = STOTOT
     '
12   TORECEP = 0
     Call ABRESTOCKS
     SQLX$ = " SELECT CANTNOM, NumProv, NUMEBORE , FECHRECEP, NumRem_Prov, CanToApro  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '
     'Set COBRECEPTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
'     Dim COBRECEPTMP As ADODB.Recordset
'     Set COBRECEPTMP = New ADODB.Recordset
'     COBRECEPTMP.CursorType = adOpenDynamic
'     COBRECEPTMP.LockType = adLockReadOnly
'     COBRECEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set COBRECEPTMP = READSET(SQLX$)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         k& = 0
         Do While (.EOF = False)
            CARECEP = XVALO(!CantNom) - XVALO(!CANTOAPRO)
            If CARECEP > 0 Then
              NCI = XVALO(!NumProv)
              REMIPRO$ = SAFETEXT$(!NumRem_Prov)
              NROBOL& = XVALO(!NUMEBORE)
              FEX = CVINT(!FECHRECEP)
              LI2TEX$ = Space$(96)
              Call REPLA(LI2TEX$, "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, rasocli$((-1) * NCI), 15, 29)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CARECEP), 4, 9, CADECI%, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
              Call REPLA(LI2TEX$, "RT." + REMIPRO$, 84, 24)
              List4(1).AddItem LI2TEX$
              TOSUMIN = TOSUMIN + CARECEP
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     '
     If TORECEP > 0.05 Then
        XXX$ = "RECEPC. " + CSTRING$(MKS$(TORECEP), 4, 11, CADECI%, 2)
        List4(0).AddItem XXX$
     End If
     GoSub PRESEDIS
     '
     Call ABREPEDCLI
     SQLX$ = " SELECT NROPED, NroClie, CanPed,  CantEnt, FechEmis, AcuDespa, AcuPrepa, FeSolEnt,Num_Prov FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     If MODODOSI% = 1 Then
          SQLX$ = SQLX$ + " AND StaDespa < 2 "
          SQLX$ = SQLX$ + "AND AcuDespa < (CanPed - 0.005) "
        Else
          SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     End If
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
'     Dim PEDTEMP As ADODB.Recordset
'     Set PEDTEMP = New ADODB.Recordset
'     PEDTEMP.CursorType = adOpenDynamic
'     PEDTEMP.LockType = adLockReadOnly
'     PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NROX& = !NROPED
           NCI = !NROCLIE
           CAPEN = !CANPED - !CANTENT
           
           SALPREP& = 0
           If MODODOSI% = 1 Then
              FEMIPED% = CVINT(!FECHEMIS)
              If FEMIPED% >= FECHANUM("07/11/05") Then
                 CAPEN = XVALO(!CANPED) - XVALO(!AcuDespa)
                 SALPREP& = XVALO(!AcuPrepa) - XVALO(!AcuDespa)
              End If
           End If
           FEX = CVINT(!FeSolEnt)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
             LI2TEX$ = Space$(156)
             Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, rasocli$(NCI), 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               OBSER$ = ""
               If FEX < HOYY Then
                 ATRAS% = DIENTRE%(FEX, HOYY)
                 OBSER$ = "Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
               End If
               If SALPREP > 0 Then
                 If OBSER$ <> "" Then OBSER$ = OBSER$ + "/ "
                 OBSER$ = OBSER$ + "Pr." + TRIM$(Str$(SALPREP&))
               End If
             Call REPLA(LI2TEX$, OBSER$, 84, 24)
             NPROV = XVALO(!Num_Prov)
             Call REPLA(LI2TEX$, SAFETEXT$(NPROV), 120, 6)
             Call REPLA(LI2TEX$, rasocli$(-1 * NPROV), 126, 24)
             
             List4(1).AddItem LI2TEX$
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     GoTo 15
     '
     '
15   Call ABRECOMPRAS
     SQLX$ = " SELECT N_ORSERV, Coef_Unids, cant_ocom, Cant_Rec, Npro_Ocom, Nume_Ocom, Fentre_Sol,MONE_COS, PRUN_NETO"
     SQLX$ = SQLX$ + " FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     'Set OcoTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim OcoTemp As ADODB.Recordset
'     Set OcoTemp = New ADODB.Recordset
'     OcoTemp.CursorType = adOpenDynamic
'     OcoTemp.LockType = adLockReadOnly
'     OcoTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set OcoTemp = READSET(SQLX$)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_ORSERV
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !cant_ocom * COEFCAN
             CAREC = !Cant_Rec * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI = !Npro_Ocom
               NROX& = !nume_ocom
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   FEX = HOYY
                 End If
               LI2TEX$ = Space$(156)
               CostoCompraDolar$ = FORMATNUM$(XVALO(!prun_neto), "F10.2")
               If XVALO(!mone_cos) = 1 Then
                  CONDI$ = "NUME_OCOM = " & NROX&
                  TipoCambioComp = XVALO(VALOBADA("OCOMENCA", "COTIZA_DOLAR", CONDI$, "NOMESS"))
                  CostoCompraDolar$ = FORMATNUM$(TipoCambioComp * XVALO(!prun_neto), "F10.2")
               End If
               '
               Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
               Call REPLA(LI2TEX$, rasocli$((-1) * NCI), 15, 29)
               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               Call REPLA(LI2TEX$, OBSER$, 84, 24)
               Call REPLA(LI2TEX$, CostoCompraDolar$, 110, 10)
               List4(1).AddItem LI2TEX$
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '
     '\\\OT-06-0149-WAR-25/04/06///
22      Call ABRECOMPRAS
     SQLX$ = " SELECT Cant_Aten, Acuen_Aten, Npro_Aten, Nume_Aten, Fentre_Sol FROM AUTOENT WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND Stat_Aten < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
     '
     'Set AtenTemp = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim AtenTemp As ADODB.Recordset
'     Set AtenTemp = New ADODB.Recordset
'     AtenTemp.CursorType = adOpenDynamic
'     AtenTemp.LockType = adLockReadOnly
'     AtenTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     '
     Set AtenTemp = READSET(SQLX$)
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = 0 '!N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = 1 '!Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !Cant_Aten * COEFCAN
             CAREC = !Acuen_Aten * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI = !Npro_Aten
               NROX& = !Nume_Aten
               FEX = CVINT(!Fentre_Sol)
               FESOL = CVINT(!Fentre_Sol)
               FEX = FESOL: OBSER$ = ""
                 If FEX < HOYY Then
                   ATRAS% = DIENTRE%(FESOL, HOYY)
                   OBSER$ = Left$(FECHATEX$(FESOL), 5) + " - Atr." + TRIM$(Str$(ATRAS%)) + " Ds."
                   FEX = HOYY
                 End If
               LI2TEX$ = Space$(156)
               Call REPLA(LI2TEX$, "AE." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
               Call REPLA(LI2TEX$, rasocli$((-1) * NCI), 15, 29)
               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 44, 9)
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               Call REPLA(LI2TEX$, OBSER$, 84, 24)
               List4(1).AddItem LI2TEX$
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '\\\OT-06-0149-WAR-FIN///
     '
25   Call ABREORFAB
     SQLX$ = " SELECT CanProy,  CanApro,  NUMEORFA, FechEnSol, Referen, FechProyEntre, FechRecal FROM ORDEFABR WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     'Set OfaTemp = Ofabric.OpenRecordset(sqlx$, dbOpenSnapshot)
'     Dim OfaTemp As ADODB.Recordset
'     Set OfaTemp = New ADODB.Recordset
'     OfaTemp.CursorType = adOpenDynamic
'     OfaTemp.LockType = adLockReadOnly
'     OfaTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set OfaTemp = READSET(SQLX$)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CACOM = !CANPROY
           CAREC = !CanApro
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             NROPED& = !NUMEORFA
             FEX = CVINT(!FechEnSol)
             REFE$ = !Referen
             LI2TEX$ = Space$(156)
             FPE% = CVINT(!FechProyEntre)
             Call REPLA(LI2TEX$, "OF." + TRIM$(CSTRING$(MKS$(NROPED&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, REFE$, 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 44, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
             'Call REPLA(LI2TEX$, IDENSEDE$, 84, 8)
             '
             If FPE% > 0 Then
               FEX = FPE%
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
               FEX = CVINT(!FechRecal)
               Call REPLA(LI2TEX$, "Calc." + FECHATEX$(FEX), 84, 24)
             End If
             '
             List4(1).AddItem LI2TEX$
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     'HASTA ACA ACTUALIZADO.
35   Call ABRESERVA
     SQLX$ = " SELECT CantRes, CANTCONS, Idsubor, IdenLote,  FechRes FROM RESERVA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
     '
     'Set ReserTemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
     
'     Dim ReserTemp As ADODB.Recordset
'     Set ReserTemp = New ADODB.Recordset
'     ReserTemp.CursorType = adOpenDynamic
'     ReserTemp.LockType = adLockReadOnly
'     ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     '
     Set ReserTemp = READSET(SQLX$)
     '
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           TORE# = XVALO(!CANTRES)
           CACO# = XVALO(!CANTCONS)
           '
           If TRIM$(UCase$(Unimed$(CIXI%))) = "GK" Then
              TORE# = TORE# / 1000
              CACO# = CACO# / 1000
           End If
           '
           SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
           SARE = TORE# - CACO#
           If SARE * SIGNO% < 0 Then
             SARE = 0
           End If
           CAPEN = SARE
           If Abs(CAPEN) > 0 Then
             NORFA$ = TRIM$(SAFETEXT$(!idsubor))
             IDLOTXX$ = TRIM$(SAFETEXT$(!idenlote))
             OBSER$ = ""
             If IDLOTXX$ <> "" Then OBSER$ = "Lote " + IDLOTXX$
             FEX = CVINT(!FechRes)
             If Left$(NORFA$, 2) = "OR" Then
               REFE$ = "Reserva p/Reparación"
               GoTo 38
             End If
             '
             DESTIN$ = NORFA$
             CICON% = CODIFAB%(NORFA$)
             If CICON% > 0 Then
               DESTIN$ = TRIM$(CODEXT$(CICON%)) + " - " + TRIM$(DESCRIT0$(CICON%))
             End If
             '
             REFE$ = "Reserva p/" + DESTIN$
             '
38            LI2TEX$ = Space$(156)
             Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
             Call REPLA(LI2TEX$, REFE$, 15, 29)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
             Call REPLA(LI2TEX$, OBSER$, 84, 24)
             List4(1).AddItem LI2TEX$
             TORESER = TORESER + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     GoSub PRESEDIS
     '
     STOTER = 0
     Call ABRESTOCKS
     SQLX$ = "SELECT CANT_ENTREG, CANT_RENDID, Cant_Devu, CANT_ASIG, Nro_OCOM, NREMI_CONSIG, Npro_Ocom, FECHA_ENTREG, Femi_Ocom FROM MACONSIG WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND (CANT_ENTREG-CANT_RENDID > 0 OR CANT_ASIG > CANT_ENTREG OR CANT_DEVU <> 0) "
     SQLX$ = SQLX$ + "AND NPRO_OCOM > 0 "
     SQLX$ = SQLX$ + " ORDER BY Femi_Ocom "
     'Set LIMACONTMP = Stocks.OpenRecordset(sqlx$, 4)
     '
'     Dim LIMACONTMP As ADODB.Recordset
'     Set LIMACONTMP = New ADODB.Recordset
     'LIMACONTMP.CursorType = adOpenDynamic
     'LIMACONTMP.LockType = adLockReadOnly
     'LIMACONTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
     Set LIMACONTMP = READSET(SQLX$)
     With LIMACONTMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TORE# = XVALO(!cant_entreg)
           CACO# = XVALO(!CANT_RENDID)
           DEVU# = XVALO(!cant_devu)
           CAPEN = TORE# - CACO# - DEVU#
           CAPENENT# = XVALO(!cant_asig) - XVALO(!cant_entreg)
           NUOC& = XVALO(Mid$(!nro_ocom, 4, 6))
           '
           If Abs(CAPEN) > 0 Then
              NUREX$ = "RC-" + TRIM$(!NREMI_CONSIG)
              NORFA$ = SAFETEXT$(!nro_ocom)
              NPRO% = XVALO(!Npro_Ocom)
              REFE$ = rasocli$((-1) * NPRO%)
              FEX = CVINT(!FECHA_ENTREG)
              SDOR$ = ""
              LI2TEX$ = Space$(156)
              '
              Call REPLA(LI2TEX$, Left$(NUREX$, 12), 1, 12)
              Call REPLA(LI2TEX$, REFE$, 15, 30)
              If CAPEN > 0 Then
                   Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, CADECI%, 2), 44, 9)
                 Else
                   Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPEN)), 4, 9, CADECI%, 2), 53, 9)
              End If
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
              Call REPLA(LI2TEX$, NORFA$, 84, 24)
              List4(1).AddItem LI2TEX$
              STOREMO = STOREMO + CAPEN
              STOTER = STOTER + CAPEN
              GoTo 47
           End If
           '
47         If CAPENENT# >= 0.05 Then
             If (STATOCOM%(NUOC&) < 8 And STATOCOM%(NUOC&) >= 0) Then
                NORFA$ = SAFETEXT$(!nro_ocom)
                NPRO% = XVALO(!Npro_Ocom)
                REFE$ = rasocli$((-1) * NPRO%)
                FEX = CVINT(!Femi_Ocom)
                SDOR$ = ""
                LI2TEX$ = Space$(156)
                '
                Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
                Call REPLA(LI2TEX$, REFE$, 15, 30)
                If CAPENENT# > 0 Then
                     Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPENENT#), 4, 9, CADECI%, 2), 53, 9)
                   Else
                     Call REPLA(LI2TEX$, CSTRING$(MKS$(Abs(CAPENENT#)), 4, 9, CADECI%, 2), 44, 9)
                End If
                Call REPLA(LI2TEX$, FECHATEX$(FEX), 74, 9)
                List4(1).AddItem LI2TEX$
                TORESER = TORESER + CAPENENT
             End If
           End If
           '
         .MoveNext
         Loop
       End If
     End With
     '
     If Abs(STOTER) >= 0.005 Then
        XXX$ = "TERCEROS" + CSTRING$(MKS$(STOTER), 4, 11, CADECI%, 2)
        List4(0).AddItem XXX$
     End If
     '
     GoSub PRESEDIS
     '
     ' ORDENAMIENTO POR FECHA
     For U& = 1 To List4(1).ListCount
45     TTXX$ = List4(1).List(U& - 1)
       FE1% = FECHANUM(Mid$(TTXX$, 74, 8))
       If Mid$(UCase$(TTXX$), 15, 12) = "STOCK PROPIO" Then FE1% = 0
       For U2& = U& + 1 To List4(1).ListCount
         TTYY$ = List4(1).List(U2& - 1)
         FE2% = FECHANUM(Mid$(TTYY$, 74, 8))
         If FE2% < FE1% Then
           List4(1).RemoveItem U2& - 1
           List4(1).AddItem TTYY$, U& - 1
           GoTo 45
         End If
       Next U2&
     Next U&
     List4(1).Refresh
     
     '
     ' STOCK PROYECTADO
     STOPRO = 0
     For U& = 1 To List4(1).ListCount
       TTXX$ = List4(1).List(U& - 1)
       STOPRO = STOPRO + XVALO(Mid$(TTXX$, 44, 9)) - XVALO(Mid$(TTXX$, 53, 9))
       'STOPROX$ = FORMATNUM$(STOPRO, "F10.1")
       STOPROX$ = FORMATNUM$(STOPRO, "F10." & CNTDC$(CIXI%))
       Call REPLA(TTXX$, STOPROX$, 62, 10)
       List4(1).RemoveItem U& - 1
       List4(1).AddItem TTXX$, U& - 1
     Next U&
     List4(1).Refresh
     '
     Call PASAR_LIST_a_Grilla(Me.List4(1), MSF(0), "1,12;15,30;44,9;53,9; 62,10;74,9;84,24;110,10;120,6;126,20")
'20   LIUBIC.Clear
'     RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CIXI%))
'     For U& = 1 To Len(RGXX$) Step 4
'        REG& = CVS(Mid$(RGXX$, U&, 4))
'        X$ = REGLEIDO$("UBISTOCK", REG&)
'        If CVI(Left$(X$, 2)) <> CIXI% Then
'           X$ = REGLEIDO$("UBISTOCK", 1)
'           Call GRAREG("UBISTOCK", X$, 1)
'           Call CIERRARCH("UBISTOCK")
'           GoTo 20
'        End If
'        '
'        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
'           Y$ = Mid$(X$, 3, 24)
'           LIUBIC.AddItem Y$
'        End If
'     Next U&
     '
     Screen.MousePointer = 1
'     Call GRAPHDIS
     '
     Exit Sub
     '
PRESEDIS:
     '
     
'     STOMI.Caption = TRIM$(CSTRING$(MKS$(STOMINI), 4, 12, CADECI%, 2))
'     LECONOMI.Caption = TRIM$(CSTRING$(MKS$(LECONOM), 4, 12, CADECI%, 2))
     '
     Label15(210).Caption = TRIM$(CSTRING$(MKS$(STOPROP), 4, 12, CADECI%, 2))
     Label15(211).Caption = TRIM$(CSTRING$(MKS$(STOREMO), 4, 12, CADECI%, 2))
     STOTOT = STOPROP + STOREMO
     Label15(212).Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, CADECI%, 2))
     '
     Label15(213).Caption = TRIM$(CSTRING$(MKS$(TOPEDID), 4, 12, CADECI%, 2))
     Label15(214).Caption = TRIM$(CSTRING$(MKS$(TORESER), 4, 12, CADECI%, 2))
     TOCOMPR = TOPEDID + TORESER
     Label15(215).Caption = TRIM$(CSTRING$(MKS$(TOCOMPR), 4, 12, CADECI%, 2))
     '
     Label15(216).Caption = TRIM$(CSTRING$(MKS$(TOSUMIN), 4, 12, CADECI%, 2))
     Label15(217).Caption = TRIM$(CSTRING$(MKS$(TOFABRI), 4, 12, CADECI%, 2))
     TOINPRE = TOSUMIN + TOFABRI
     Label15(218).Caption = TRIM$(CSTRING$(MKS$(TOINPRE), 4, 12, CADECI%, 2))
     '
     STONET = STOTOT - TOCOMPR
     STODIS = STONET - STOMINI + TOINPRE
     STOFAL = 0
     If STODIS < 0 Then
        STOFAL = Abs(STODIS)
        STODIS = 0
     End If
     '
'     Label401.Caption = TRIM$(CSTRING$(MKS$(STONET), 4, 12, CADECI%, 2))
'     Label402.Caption = TRIM$(CSTRING$(MKS$(STODIS), 4, 12, CADECI%, 2))
'     Label403.Caption = TRIM$(CSTRING$(MKS$(STOFAL), 4, 12, CADECI%, 2))
'     Label404.Caption = TRIM$(CSTRING$(MKS$(XVALO(Label103) - XVALO(Label202)), 4, 12, CADECI%, 2))
     '
     DoEvents
     '
Return
'
COBRECEPTMP.Close
Set COBRECEPTMP = Nothing
PEDTEMP.Close
Set PEDTEMP = Nothing
OcoTemp.Close
Set OcoTemp = Nothing
AtenTemp.Close
Set AtenTemp = Nothing
OfaTemp.Close
Set OfaTemp = Nothing
ReserTemp.Close
Set ReserTemp = Nothing
LIMACONTMP.Close
Set LIMACONTMP = Nothing
'
Screen.MousePointer = 1
End Sub

Sub Guardar_Rangos()
  Dim FORMU As ADODB.Recordset
  Set FORMU = New ADODB.Recordset
  Ci% = CODINT%(Text1)
  Set FORMU = FLEXCONN.Execute(FILTSQL$("Select * from FormuPat where CodIConj = " & Ci% & " ORDER BY NUORIT ASC"))
  If FORMU.BOF And FORMU.EOF Then
     Exit Sub
  End If
  '
   FORMU.MoveFirst
   REG& = 0
   Do While Not (FORMU.EOF)
     AAXX$ = REGBLAN("!RACEPTA")
     CODIELE% = XVALO(FORMU!codielem)
     Unidad$ = Unimed$(CODIELE%)
     If UCase(Unidad$) = "KG" Or UCase(Unidad$) = "GR" Or UCase(Unidad$) = "MG" Or UCase(Unidad$) = "MC" Then
       Uso# = XVALO(FORMU!usobruto)
       MINIMO# = XVALO(FORMU!Acepta_Min)
       Maximo# = XVALO(FORMU!Acepta_Max)
       PREC# = XVALO(FORMU!Acepta_Prec)
       NUMITEM% = XVALO(FORMU!NUORIT)
       '
       Call REPLA(AAXX$, MKI(CODIELE%), 1, 2)
       Call REPLA(AAXX$, Unidad$, 3, 4)
       Call REPLA(AAXX$, MKD(Uso#), 7, 8)
       Call REPLA(AAXX$, MKD(MINIMO#), 15, 8)
       Call REPLA(AAXX$, MKD(Maximo#), 23, 8)
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
   CIXI% = CODINT%(Text1)
   LOTE# = LOTESTAN(CIXI%)
   TITU$ = "Rango de Aceptación:  " & Text1 & "  x  " & LOTE# & "  " & Unimed(Ci%) & "  -  " & DESCRIT0(Ci%)
   '
10 Respuesta% = VENTABU("RANGACEP/TITUPAN=" & TITU)
   If Respuesta% < 0 Then
      Exit Sub
   End If
   '
  CIXI% = CODINT%(Text1)
  '
  FIN& = ULTREG("!RACEPTA")
  For REG& = 1 To FIN&
     AARR$ = REGBLAN("!RACEPTA")
     AARR$ = REGLEIDO("!RACEPTA", REG&)
     NUMITEM% = CVI(Mid(AARR$, 39, 2))
     CODIELE% = CVI(Mid(AARR$, 1, 2))
     Unidad$ = Unimed(CODIELE%)
     Uso# = CVD(Mid(AARR$, 7, 8))
     MINIMO# = CVD(Mid(AARR$, 15, 8))
     If MINIMO# > 0 Then
        COMUNI ("El Rango de Aceptación Minimo Debe ser Menor o Igual a Cero")
        GoTo 10
     End If
     '
     Maximo# = CVD(Mid(AARR$, 23, 8))
     If Maximo# < 0 Then
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
     Call ACTUREGISTRO("FORMUPAT", "ACEPTA_MAX", CONDI$, Maximo#, REGI&)
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
     Colors$ = lColor$(CC%)
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
          !CODIEMPR = CodiEmp%
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
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !Codigo = CODIG
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
      For k% = 1 To Len(DPS$) Step 8
        DPSX$ = Mid$(DPS$, k%, 8)
        CIXXII% = CVI(Left$(DPSX$, 2))
        If CIXXII% > 0 Then
          If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
            With Formulas
              On Error Resume Next
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !CODICONJ = CICONJ%
              !CODXCONJ = CODEXT$(CICONJ%)
              !codielem = CIXXII%
              !CodXElem = CODEXT$(CIXXII%)
              !usobruto = CVS(Mid$(DPSX$, 3, 4))
              !PLADISPO = Asc(Mid$(DPSX$, 7, 1))
              .Update
              .Bookmark = .LastModified
            End With
          End If
        End If
      Next k%
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
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
         Ci% = CVI(Left$(X$, 2))
         RESECOX$ = X$
      !CODICONJ = Ci%
      !CODXCONJ = CODEXT$(Ci%)
      !NUMEOPER = CVI(Mid$(RESECOX$, 3, 2))
      !descoper = TRIM$(Mid$(RESECOX$, 5, 48))
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

Sub LIMPIARSTOCKS()
   For i% = 210 To 218
      Label15(i%) = ""
   Next i%
   MSF(0).Rows = 1
   MSF(1).Rows = 1
End Sub

Private Sub Abri_Click()
Call Command9_Click
End Sub

Private Sub Acopia_Click()
   Call Command22_Click
End Sub

Private Sub Anuevo_Click()
   Call Command21_Click
   'SELECT PARA SABER SI EN BASES MISMAS TABLAS MUESTRA LAS FILAS QUE NO MACHEAN UNA CON OTRA
'SELECT m.codint,m.codigo
'FROM mineparts.dbo.master m
'except
'SELECT mm.codint,mm.codigo
'FROM ahp.dbo.master mm

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
    Case 2
      If Command47(2).Caption = "Visualizar costos" Then
        Command47(2).Caption = "Refrescar Costos"
      Else
        Command47(2).Caption = "Visualizar costos"
      End If
      Call CARGAR_STOCKS_PROVEEDOR(XVALO(Label28), 1)
    Case 3
      If Command47(3).Caption = "Visualizar costos" Then
        Command47(3).Caption = "Refrescar Costos"
      Else
        Command47(3).Caption = "Visualizar costos"
      End If
      Call GENKARDEXAHP(1)

   End Select
   '
'   MODIFIC%(1) = 1
   '
99 Command47(Index).Enabled = True

End Sub
Sub CARGARUTA()
      'PERMITE AGREGAR NUEVOS POE'S
      '
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", DESCRI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, Ruta, 32, 512)
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
   CONDI$ = "Codint=" & CODINT%(Text1)
   Call ACTUREGISTRO("Master", "UsuApFor", CONDI$, USERNAME, REG&)
   Call ACTUREGISTRO("Master", "FechApFor", CONDI$, CVDAT(HOII%), REG&)
   Call ACTUREGISTRO("Master", "Revi_Formu", CONDI$, CVDAT(HOII%), REG&)
   Label59.Caption = FECHATEX$(FEREFOR%(Ci%))
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
   If Control$("FACTURA", "SALIDA") <> "FELECT" Then ' TIENE QUE ESTAR CONFIGURADO FACT.ELEC.
      Call MENSERR(24, "Opción Solo Para Factura Electrónica")
      GoTo 99
   End If
   '
20 OPX% = ALTERNA%("  Actualización  \  Listados  ")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   '
   If OPX% = 1 Then  ' ACTUALIZACION
      Ci% = CODINT%(Text1) ' VALIDA QUE SE HAYA SELECCIONADO CODIGO
      If Ci% < 1 Then
        Call MENSERR(24, "Código No Válido")
        GoTo 99
      End If
      '
      'TOMO LOS VALORES DE LA BASE DE DATOS PARA PRESENTAR EN EL OBJETO
      CONDI$ = "Codint = " & Ci%
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
    Command54(Index).Enabled = False

    If DERACCE%("MODIMETE", "Modificacion de Memoria Técnica") < 2 Then
        GoTo 99
    End If

    Select Case Index
    Case 0
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      Label63.Caption = Ruta$
    Case 1
      Command54(Index).Enabled = False 'MODIFICADO OD-30/11/11 NO EXISTE COMMAND71
      Label63 = ""
      Command54(Index).Enabled = True
    End Select

99  Command54(Index).Enabled = True
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
   X$ = OBJPLA$("UPDATELTC", VDEF$)
   If X$ <> "" Then
     PRO% = CVI(Mid$(X$, 1, 2))
     LTC$ = TRIM$(Mid$(X$, 3, 5))
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
       strSQL = "SELECT * FROM " & TAPRIN$
       AUX1.Open strSQL, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
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

   For i% = 0 To List1.ListCount - 1 ' ORDENAMIENTO
     CI1$ = TRIM$(Mid$(List1.List(i%), 1, 16))
     For J% = i% + 1 To List1.ListCount - 1
       CI2$ = TRIM$(Mid$(List1.List(J%), 1, 16))
       If UCase$(CI1$) > UCase$(CI2$) Then
         COD1$ = List1.List(i%)
         COD2$ = List1.List(J%)
         List1.List(i%) = COD2$
         List1.List(J%) = COD1$
         CI1$ = CI2$
       End If
     Next J%
   Next i%
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
    Ci% = CODINT%(Text1)
    COPE% = XVALO(Label55)
    '
    'VUELCO LA INFORMACION DE LA BASE A LA VENTANA
    J& = 0
    SQLX$ = "SELECT * FROM HERRAOPER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODINT =" & Ci%
    SQLX$ = SQLX$ + " AND Cod_Oper = " & COPE%
    SQLX$ = SQLX$ + " AND Tipo_Herr = " & TIPOHERRAM%
    Dim HERRAOPTMP As ADODB.Recordset
    Set HERRAOPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With HERRAOPTMP
      Do While Not .EOF
        CODIG$ = SAFETEXT$(!Cod_Herramental)
        CANTI = XVALO(!Cantidad)
        If CODIG$ <> "" Then
        Call REPLA$(X$, CODIG$, 1, 16)
        Call REPLA$(X$, MKS(CANTI), 17, 4)
        J& = J& + 1
        Call GRAREG("!HERRAOP", X$, J&)
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
    CONDI$ = "CODINT =" & Ci% & " AND  Cod_Oper = " & COPE% & " AND Tipo_Herr = " & TIPOHERRAM%
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
        X$ = REGLEIDO$("!HERRAOP", U&)
        CODIG$ = TRIM$(Left$(X$, 16))
        CANTI = CVS(Mid$(X$, 17, 4))
        If CODIG$ <> "" Then
           If CANTI < 1 Then Call COMUNI("Falta Ingresar Cantidad para Código: " & CODIG$): GoTo 20
            With HERRAOPERTMP
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CODIEMPR = CodiEmp%
                On Error GoTo 0
                !CODINT = Ci%
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then Call FINAL("")
End Sub

Private Sub Label15_Change(Index As Integer)
    If Index <> 2 And Index <> 8 And Index <> 10 Then Exit Sub
    
    If SSTab1.Tab = 4 Then
      If Label15(2) = "" Then
           Call LIMPIARSTOCKS
           MSF(0).Rows = 1
           MSF(1).Rows = 1
           Exit Sub
      End If
      CI1% = XVALO(Label28)
      Screen.MousePointer = 11
      Call CARGAR_STOCKS
      Call CARGAR_STOCKS_PROVEEDOR(CI1%)
      
      Screen.MousePointer = 1
    End If
    '
    If SSTab1.Tab = 5 Then
      If Label15(8) = "" Then
           MSF(3).Rows = 1
           MSF(4).Rows = 1
           TEPRUEBA.Font = MSF(3).Font
           Exit Sub
      End If
      CI1 = XVALO(Label28)
      Screen.MousePointer = 11
      Call GENKARDEXAHP
      Screen.MousePointer = 1
    End If
    
    If SSTab1.Tab = 6 Then
      If Label15(10) = "" Then
        MSF(5).Rows = 1: MSF(7).Rows = 1: MSF(6).Rows = 1: MSF(8).Rows = 1
        Exit Sub
      End If
      CI1 = XVALO(Label28)
      Screen.MousePointer = 11
      Call CARGA_ESTADISTICAS(CI1%)
      Screen.MousePointer = 1
    End If
    '
End Sub

Private Sub Label15_DblClick(Index As Integer)
   If (Index = 215 Or Index = 218) And TRIM$(TXTBUSCAR) <> "" Then
      TXTBUSCAR = ""
      Label15(215).BackColor = &HE0E0E0
      Label15(218).BackColor = &HE0E0E0
      Exit Sub
   End If
   
   If Index = 215 Then
      TXTBUSCAR = "PC."
      CMBFILTRO.TEXT = MSF(0).TextMatrix(0, 1)
      Label15(215).BackColor = &HC0FFFF
   ElseIf Index = 218 Then
      CMBFILTRO.TEXT = MSF(0).TextMatrix(0, 1)
      Label15(218).BackColor = &HC0FFFF
      TXTBUSCAR = "OC."
   ElseIf Index = 212 Then
       Label15(218).BackColor = &HE0E0E0
       Label15(215).BackColor = &HE0E0E0
       Call TXTBUSCAR_KeyPress(27)
   End If
   
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
  CIXI% = CODINT%(Text1)
  If CIXI% <= 0 Then
      Exit Sub
  End If
  LOTE# = LOTESTAN(CIXI%)
  If LOTE# < 1 Then
     LOTE# = 1
   End If
  TITU$ = "Form. Patron:  " & Text1 & "  x  " & LOTE# & "  " & Unimed(Ci%) & "  -  " & DESCRIT0(Ci%)
  Call SETULTREG("!MATEROF", 0)
  RESP% = COMALTER%("Opciones:\\ABM de Formula Patron \ Rangos de Aceptacion")
  If RESP% < 1 Then GoTo 99
  If RESP% = 2 Then
     Call Guardar_Rangos
     GoTo 99
  End If
  Dim FORMU As ADODB.Recordset
  Set FORMU = New ADODB.Recordset
  Set FORMU = FLEXCONN.Execute(FILTSQL$("Select * from FormuPat where CodIConj = " & Ci% & " ORDER BY NUORIT ASC"))
  
  If FORMU.BOF And FORMU.EOF Then
     GoTo 10
  End If
   
  'Completa MATEROF para presentar en pantalla
     FORMU.MoveFirst
     REG& = 0
   Do While Not (FORMU.EOF)
     AAXX$ = REGBLAN("MATEROF")
     CODIELE% = XVALO(FORMU!codielem)
     Unidad$ = Unimed$(CODIELE%)
     Uso# = XVALO(FORMU!usobruto)
     Operacion% = XVALO(FORMU!Operacion)
     '
     Call REPLA(AAXX$, MKI(CODIELE%), 1, 2)
     Call REPLA(AAXX$, Unidad$, 3, 2)
     Call REPLA(AAXX$, MKD(Uso#), 9, 8)
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
     If Control("FORMULAS", "REQAPRO") = "SI" Then
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
     CONDI$ = "CodIConj = " & Ci%
     Call ACTUREGISTRO("FORMUPAT", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
       
     'Recorre archivo Materof para actualizar Tabla Formupat
     FORMU.Close
     FORMU.Open FILTSQL$("Select * from formupat where CodIConj = " & Ci%), FLEXCONN, adOpenKeyset, adLockPessimistic
     REGACTI& = 1
     REGFINAL% = ULTREG("!Materof")
     Do While REGACTI& <= REGFINAL%
          AAXX$ = REGLEIDO("!MATEROF", REGACTI&)
          CODIELE% = CVI(Mid$(AAXX$, 1, 2))
          Unidad$ = Mid$(AAXX$, 3, 2)
          Uso# = CVD(Mid$(AAXX$, 9, 8))
          OPERA% = CVI(Mid$(AAXX$, 29, 2))
          '
          FORMU.AddNew   ' YA CONVERTIDO
               On Error Resume Next
          FORMU!CODIEMPR = CodiEmp%
               On Error GoTo 0
           FORMU!CODICONJ = CIXI%
           FORMU!CODXCONJ = CODEXT$(CIXI%)
           FORMU!codielem = CODIELE%
           FORMU!CodXElem = CODEXT$(CODIELE%)
           FORMU!usobruto = Uso#
           FORMU!Operacion = OPERA%
           FORMU!NUORIT = REGACTI&
           FORMU!MarBorra = 0
          FORMU.Update
          
          REGACTI& = REGACTI& + 1
      Loop
      FORMU.Close
      Set FORMU = Nothing
      '
      CONDI$ = "CodIConj = " & Ci% & " AND MARBORRA = 1 "
      Call BORRAREGISTROS("FORMUPAT", CONDI$, RAFE&, "NOMESS")
      '
      ' PONE MARCA DE BORRADO A PRE-EXISTENTES EN FORMULAS
      CONDI$ = "CodIConj = " & Ci%
      Call ACTUREGISTRO("FORMULAS", "MARBORRA", CONDI$, 1, RAFE&, "NOMESS")
      '
      Dim FORMUX As ADODB.Recordset
      Set FORMUX = New ADODB.Recordset
      FORMUX.Open FILTSQL$("select * from Formulas where codiconj = " & CIXI%), FLEXCONN, adOpenKeyset, adLockPessimistic
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
            Uso# = ((XVALO(FORPAT!SUBRUTO)) / LOTE#)
            CODIELE% = FORPAT!codielem
            FORMUX.AddNew   ' YA CONVERTIDO
                  On Error Resume Next
            FORMUX!CODIEMPR = CodiEmp%
                  On Error GoTo 0
            FORMUX!CODICONJ = CIXI%
            FORMUX!CODXCONJ = CODEXT$(CIXI%)
            FORMUX!codielem = CODIELE%
            FORMUX!CodXElem = CODEXT$(CODIELE%)
            FORMUX!usobruto = Uso#
            FORMUX!NUORIT = NumOrd%
            FORMUX!MarBorra = 0
            FORMUX.Update
          FORPAT.MoveNext
          Loop
      End If
      FORMUX.Close
      Set FORMUX = Nothing
      '
      CONDI$ = "CodIConj = " & CIXI% & " AND MARBORRA = 1 "
      Call BORRAREGISTROS("FORMULAS", CONDI$, RAFE&, "NOMESS")
      
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "CodIConj = " & CIXI%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "FORMULAS", "MINEPARTS", "FORMULAS", CONDI$)
      End If
     
     Call Text1_Change
99
End Sub



Private Sub Label28_Change()
   If SSTab1.Tab = 4 Then
      Label15(2) = Label28
   End If
   If SSTab1.Tab = 5 Then
      Label15(8) = Label28
   End If
   If SSTab1.Tab = 6 Then
      Label15(10) = Label28
   End If

End Sub

Private Sub Label3_DblClick()
    
    If DERACCE%("MODIMETE", "Modificacion de Memoria Técnica") < 1 Then
        Exit Sub
    End If

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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then
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
   DOPEP$ = DESCRIOP$(Ci%, NOPEP%)
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
30 RESECO$(ILISTAX&) = Left$(MKI$(Ci%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(Ci%) + MKI$(NOPY%) + String$(40, 0) 'para acomodar hoja de procesos DE FABRICACION
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

Sub Command6_Click()
   Command6.Enabled = False
50  TTOP$ = ""
    OPMAX% = 0
    OPTRAZA% = 0
    OPNOVE% = 0
    OPREPORTE% = 0
    CARCICLO% = 0
    '
    'arma la lista del alterna
    If EXISTE%("TRAZAB07.EXE") > 0 Then
       TTOP$ = "Costo de Reposición por Lote"
       OPTRAZA% = 1
       OPMAX% = 1
    End If
    OPMAX% = OPMAX% + 1
    If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
    OPNOVE% = OPMAX%
    TTOP$ = TTOP$ + "Actualización de Costos Por Proveedor"
    '
    EPAC$ = TRIM$(EMPREAC$)
    If InStr(EPAC$, "ALEJPA") Then
      OPMAX% = OPMAX% + 1
      If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
      CARCICLO% = OPMAX%
      TTOP$ = TTOP$ + "Cargar Ciclos"
    End If
    
    OPMAX% = OPMAX% + 1
    If OPMAX% > 1 Then TTOP$ = TTOP$ + "\"
    OPREPORTE% = OPMAX%
    TTOP$ = TTOP$ + "Estadísticas de Cambios de Precios"
'    '
    If TTOP$ <> "" Then
       OPMAX% = OPMAX% + 1
       TTOP$ = TTOP$ + "\Otros Accesos Directos"
    End If
    '
    'fin armado lista de alterna
    '
    'segun opcion elegida muestra o ejecuta
    If OPMAX% < 1 Then GoTo 99
    '
    OPX% = ALTERNA%(TTOP$)
    If OPX% < 1 Or OPX% > OPMAX% Then GoTo 99

    If OPX% = OPTRAZA% Then
        ACOSLOTE.Show 1
        GoTo 50
    ElseIf OPX% = OPNOVE% Then
        If DERACCE%("NOVECOST", "Ver Cambios de Costos") = 2 Then
            Call ACTUCOSTO_PROVEEDOR
            GoTo 50
        End If
    ElseIf OPX% = CARCICLO% Then
        Call CARGA_CICLOS
        GoTo 50
    ElseIf OPX% = OPREPORTE% Then
        If DERACCE%("NOVECOST", "Ver Cambios de Costos") = 2 Then
            Call ESTAXPRV
            GoTo 50
        End If
    Else
        Call ACCDIR("AMASTOK")
        GoTo 50
    End If
   '

'   If EXISTE%("TRAZAB07.EXE") Then
'        OPX% = ALTERNA%("Costo de Reposición por Lote\Otros Accesos Directos")
'        If OPX% = 1 Then
'            ACOSLOTE.Show 1
'          ElseIf OPX% = 2 Then
'            Call ACCDIR("AMASTOK")
'        End If
'        GoTo 99
'   End If
'   '
'   EPAC$ = TRIM$(EMPREAC$)
'   If InStr(EPAC$, "ALEJPA") Then
'      Call CARGA_CICLOS
'   Else
'      Call ACCDIR("AMASTOK")
'
'   End If

99 Command6.Enabled = True
End Sub

   Sub CARGA_ENCABEZADO_X_ART()
   
     With INFO_GRILLA07
        .MSF.Row = 1
        .MSF.Cols = 10
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        '
        NUCOLU% = 2: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha"
        NUCOLU% = 3: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 4: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 5: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        '
        NUCOLU% = 6: .TEPRUEBA = "123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        NUCOLU% = 7: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Usuario"
        NUCOLU% = 8: .TEPRUEBA = "1234567890123456789+123456789+123456789+123456789+123456789+": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Observación"
        '
        NUCOLU% = 9: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "M.Original"
        NUCOLU% = 10: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cotiz."
       
     End With
End Sub

Sub CARGA_CICLOS()
    '
    Call CREACAMPO("", "MASTER", "CANT_PZA_CICLO", 4, 0)
    Call CREACAMPO("", "MASTER", "CANT_EQUIV_SEGUNDOS", 4, 0)
    Call CREACAMPO("", "MASTER", "TMPO_CICLO_REAL", 4, 0)
    '
    'VALIDO QUE HAYA UN CODIGO ACTIVO
    Ci% = XVALO(Label28)
    If Ci% < 1 Then
      Call COMUNI("Código Inexistente o No Válido")
      Exit Sub
    End If
    '
    CONDI$ = "CODINT = " & Ci%
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
       Y$ = REGBLAN$("ACPREVEN")
       CLIS$ = TRIM$(Str$(U&))
       While Len(CLIS$) < 3: CLIS$ = "0" + CLIS$: Wend
       APREC$ = "LIPRE" + CLIS$
       PRLST# = PRELISTA#(APREC$, CIXX%)
       If PRLST# >= 0.00005 Then
          Call REPLA(Y$, MKD$(PRLST#), 33, 8)
       End If
       Call REPLA(Y$, NLIST$, 1, 32)
       J& = J& + 1
       Call GRAREG("!ACPREVEN", Y$, J&)
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
        Y$ = REGLEIDO$("!ACPREVEN", NUMLIS&)
        PRLST# = CVD(Mid$(Y$, 33, 8))
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
           !CODIEMPR = CodiEmp%
           On Error GoTo 0
           !Cod_Lista = NUMLIS&
           !Cod_Pieza = CIXX%
           !Num_Orden = MAORD& + 1
           !Desc_Pieza = Mid$(Y$, 3, 42)
           !Denom_Pieza = DESCRIT0$(CIXX%)
           !Carac_Control = Asc(Mid$(Y$, 45, 1))
           !Num_Original = Asc(Mid$(Y$, 46, 16))
           !Precio_Unitario = PRLST#             'CVD(Mid$(Y$, 62, 8))
           !Cant_Limite = CVD(Mid$(Y$, 70, 2))
           !Porc_Limite = CVD(Mid$(Y$, 72, 1))
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
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(1, EPAC$, "AHP") = 0 Then Image3.Picture = LoadPicture("")
    
    Campos$ = "Comprob./A12;Referencia/A14;Entra/F8.1;Sale/F8.1;St.Proy./F8.1;Fecha/D8;Oboservaciones/A20;Compra U$S/F10.1;Cta.Prov./f7;R.Social/A16"
    Call CARGA_ENCABEZADO(MSF(0), Campos$, AMASTOAHP, 1)
    Call CARGA_ENCABEZADO(MSF(2), Campos$, Me, 1)
 
 
    Campos$ = ("Fecha/D6;CM/A1;Transacción/A15;Doc.Número/A15;ID-Bulto/A9;Referencia/A15;Dp/F2;Entrada/F5;Salida/F5;Saldo/F5,Fob/F9.2;Nacionalizado/F12.2;P.Vta/F0;Prov./A5;V.Ent./A12")
    Call CARGA_ENCABEZADO(MSF(3), Campos$, AMASTOAHP, 1)
    Call CARGA_ENCABEZADO(MSF(4), Campos$, Me, 1)
    '
    Campos$ = "Cta./F4;R.Social/A10;V.Entrega/A10;Costo/F9.2;Nacionalizado/F10.2; U$S x Frmla./F9.2;U$S Venta/F9.2;VIENTREGANUM/F0"
    Call CARGA_ENCABEZADO(MSF(10), Campos$, AMASTOAHP)
    '
    SSTab1.Tab = 0
    SSTab1.TabVisible(2) = False
    Label10(2) = Control$("AMASTO", "ATRIPAR1")
      If Label10(2) = "" Then Label10(2) = ". Atributo 'A':"
      If Left$(Label10(2), 1) = "." Then Label10(2) = " " + Mid$(Label10(2), 2)
    Label10(3) = Control$("AMASTO", "ATRIPAR2")
      If Label10(3) = "" Then Label10(3) = ". Atributo 'B':"
      If Left$(Label10(3), 1) = "." Then Label10(3) = " " + Mid$(Label10(3), 2)
      '
    '\\\OT-05-0633-WAR-22/09/05///
    PORLARGO% = 0
    Label10(20).Visible = False
    Label10(27).Visible = False
    Text41.Visible = False
    Label10(26).Visible = False
    Label42.Visible = False
    Label10(24).Visible = False
    Label10(25).Visible = False
    Label36.Visible = False
    Text42.Visible = False
    RTA1$ = Control$("AMASTO", "FORMULA")
    If UCase$(TRIM$(RTA1$)) = "CANTIDAD" Then
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
      Label10(20).Visible = True
      Label10(27).Visible = True
      Text41.Visible = True
      Label10(26).Visible = True
      Label42.Visible = True
      Label10(24).Visible = True
      Label10(25).Visible = True
      Label36.Visible = True
      Text42.Visible = True
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
    'LLAMADA DE ANACOS09 PARA QUE LEVANTE SOLAPA DE ESPECIFICACIONES
    If Left$(APLINVO$, 8) = "P.O.E.'S" Then
        CODX$ = TRIM$(Mid$(APLINVO$, 9))
        CODX$ = REPLACAR$(CODX$, "|", "/")
        If CODINT%(CODX$) > 0 Then
            Text1 = CODX$
            Call CARESPECI
            SSTab1.Tab = 3
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
      Call Command21_Click
      Call Command3_Click
    End If
    '
    CONDI$ = "CODINT=0"
    If CantRegistros("MASTER", CONDI$) > 0 Then Call BORRAREGISTROS("MASTER", CONDI$, REG&)
    CONDI$ = "COD_INTE=0"
    If CantRegistros("LOTINGRE", CONDI$) > 0 Then Call BORRAREGISTROS("LOTINGRE", CONDI$, REG&)
    '
    If Control$("", "IDSECUN") = "ALIAS" Then
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
    Call CREATABLA_COSTOXPROVE

    HABIPORMI% = 0
    If Control$("AMASTO", "PORMILES") = "SI" Then
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
    LOMA% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LOMA% > 15 Then Text1.MaxLength = LOMA%
    '
    'Verifica si utiliza descripcion larga o descripcion de codigo y alias
    If Control("MASTER", "DESLARGA") <> "SI" Then
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
    If Control("MASTER", "FORPATRO") = "SI" Then
      PORPATRON% = 1
      Label27.ToolTipText = "Doble Clic Para Ver Formula Patron"
      Visible = True
      Screen.MousePointer = 11
      Call ACTUFECHAPRFOR
      Screen.MousePointer = 1
    End If
    '
    If Control("MASTER", "M2PORUN") = "SI" Then
        Label10(4) = "M2/ Unid."
    End If
    '
    If DERACCE%("MODIMETE/NOMESS", "Modificacion de Memoria Técnica") < 1 Then
      Label3.Visible = True
    End If
    '
    Text6(1) = "01/01/80"
    FEINIKAR$ = TRIM$(Control("", "FEINIKAR"))
    FF% = FECHANUM(FEINIKAR$)
    If FF% > 0 Then Text6(1) = FECHATEX$(FF%)
    HOII% = HOY(HOS$)
    Text6(2) = HOS$

    Call APERBASDAT("MASTER")
    '
    Call CREAR_STORED_PROCEDURE
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$    'Realiza la llamada al objeto segun la opcion seleccionada
       Case "CONFIGURA"
           Call AMASTOK07.Command37_Click       'Configuración General Maestro Items Stock
       Case "TATIMAT"
           Call OPMAST07.Option1_DblClick(0)    'Tabla de Tipo de Material
       Case "TAGRUCOVEN"
           Call OPMAST07.Option1_DblClick(3)     'Tabla Grupos Contables Ventas
       Case "TAGRUVENIVA"
           Call OPMAST07.Option1_DblClick(4)    'Tabla Grupos Ventas IVA
       Case "TAUNIMED"
           Call OPMAST07.Option1_DblClick(9)    'Tabla Unidades de Medida
       Case "TADEPOSI"
           Call OPMAST07.Option1_DblClick(5)    'Tabla de Depositos y Almacenes
       Case "TASECENCO"
           Call OPMAST07.Option1_DblClick(6)    'Tabla de Secciones y Centros de Costos
       Case "TAMAQEQUI"
           Call OPMAST07.Option1_DblClick(7)    'Tabla de Maquinas y Equipos de Producción
       Case "TACATEMPL"
           Call OPMAST07.Option1_DblClick(8)    'Tabla de Categorias de Empleados Productivos
       Case "CONLISTA"
           Call AMASTOK07.Command38_Click       'Consultas, Listados y Complementos
       Case "LISATRIGE"
           Call OPLIMAS07.Command9_Click        'Listado de Atributos Generales
       Case "STMINLOREP"
           Call OPLIMAS07.Command7_Click        'Stock Minimo y lote de Reposición
       Case "LISCAPINV"
           Call OPLIMAS07.Command15_Click       'Listado de Captura Para Inventarios
       Case "COSREPROV"
           Call OPLIMAS07.Command5_Click        'Listado de Costo de Reposicion por Proveedor
       Case "COSREPART"
           Call OPLIMAS07.Command4_Click        'Listado de Costo de Reposición por Articulo
       Case "RESTITEM"
           Call OPLIMAS07.Command3_Click        'Consulta y Restauracion de Items de Baja
       Case "ITEMINAC"
           Call OPLIMAS07.Command2_Click        'Listado de Items Inactivos
       Case "MULTINIVEL"
           Call AMASTOK07.Command40_Click(4)    'Muestra estructura Multinivel
       Case "IMPLOSION"
           Call AMASTOK07.Command40_Click(6)    'Conjuntos a los que Pertenece Item Activo
       Case "LISCOMPRA"
           Call AMASTOK07.Command40_Click(2)    'Lista de Compras
       Case "LISMULTINI"
           Call AMASTOK07.Command40_Click(1)    'Lista de Materiales Multinivel
       Case "LISENLACE"
           Call AMASTOK07.Command40_Click(0)    'Lista de Estructuras y Enlaces
       Case "LISECOPER"
           Call AMASTOK07.Command42_Click       'Listado de Secuencias de Operaciones
       Case "ALTACODIGO"
           TERMINA% = 0
           Call AMASTOK07.Command21_Click       'ALTA DE NUEVO CODIGO
       Case Else
           TERMINA% = 0
   End Select
End Sub

Private Sub mnuinformeCostosVsVentas_Click()

     Dim diccionario As Object
     Set diccionario = CreateObject("Scripting.Dictionary")
     Dim VALORES As Variant
     OPX% = COMALTER%("Seleccione Tipo de Informe\\Facturas con Detalle\Facturas Acumuladas")
     If OPX% < 1 Or OPX% > 2 Then
       GoTo 99
     End If
     
    Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
    If PERIO$ = "" Then
       GoTo 99
    End If
    
    Campos$ = "Fecha/D8"
    Campos$ = Campos$ + ";Comprobante/A18"
    Campos$ = Campos$ + ";Cta/F7"
    Campos$ = Campos$ + ";R.Social/A24"
    If OPX% = 1 Then 'detallle
      Campos$ = Campos$ + ";Codigo/A16"
      Campos$ = Campos$ + ";Lote/A12" '6
      Campos$ = Campos$ + ";Cantidad/F6" '7

    Else 'acumulado algunos campos no muestra
      Campos$ = Campos$ + ";Codigo/A0"
      Campos$ = Campos$ + ";Lote/A0"
      Campos$ = Campos$ + ";Cantidad/A0" '7
    End If
    
    Campos$ = Campos$ + ";Tot.Fact.$/F12.2" '8
   
    If OPX% = 1 Then  'detallle
        Campos$ = Campos$ + ";Tot.Fact.U$s/F12.2" '9
        Campos$ = Campos$ + ";Pr.Unit.U$s/F12.2" '10
        Campos$ = Campos$ + ";Importe.Items U$s(*1)/F14.2" '11
    Else 'acumulado algunos campos no muestra
        Campos$ = Campos$ + ";Tot.Fact.U$s(*1)/F12.2" '9
        Campos$ = Campos$ + ";Pr.Unit.U$s/A0" '10
        Campos$ = Campos$ + ";Importe.Items U$s/A0" '11
    End If
    Campos$ = Campos$ + ";Imp.Fob U$s/F12.2" '12
    Campos$ = Campos$ + ";Imp.Nacionalizado.U$s(*2)/F20.2" '13
'    Campos$ = Campos$ + ";% *1 vs. *2/F12.2" '14
    Campos$ = Campos$ + ";% *1 vs. *2/A0" '14
    Campos$ = Campos$ + ";Cotiz.U$s/F12.2" '15
    
    'agregado segun solicitud de Fernando resultado bruto y rentabilidad sobre ventas
    Campos$ = Campos$ + ";Res.Bruto/F12.2" '16  11 - 13
    Campos$ = Campos$ + ";Rent.Ventas/F12.2" '17


    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)

    DESDA = FETEXCOR1$(DESDEX%)
    HASDA = FETEXCOR1$(HASTAX%)
    'CONSULTA SIEMPRE POR DETALLE SI SE ELIGE POR DETALLE MUESTRA EL RESULTADO DE ESTA CONSULTA
    'SI NO MUESTRA EL DICCIONARIO QUE ARMA DE ACUERDO A ESTA CONSULTA
    SQLX$ = "SELECT C.TipoMovim, C.NuClien, C.NUMECOMP, C.ValAbsComp, C.OpciMovim, C.codicom_lar,C.COTIZA_DOLAR,C.idebecomp"
    SQLX$ = SQLX$ + " ,C.Mone_Emis, C.Tipo_Cam, C.FECHCONTA, C.ModoOpciMovim, DEUDOR12.RASO_CLI, DEUDOR12.VEND_CLI "
    SQLX$ = SQLX$ + " , M.COD_EXTE,M.COD_INTE,M.DESCRIP,M.VALOUNIT,M.VALTOTIT,M.IDENLOTE,M.MONEVAL,M.NUMEORDPED"
    SQLX$ = SQLX$ + " ,M.cantsalid,M.CANTINGRE"
    SQLX$ = SQLX$ + " FROM CAJABANC C WITH(NOLOCK) "
    SQLX$ = SQLX$ + " INNER JOIN DEUDOR12 WITH(NOLOCK) ON C.NUCLIEN=DEUDOR12.NUME_CLI "
    SQLX$ = SQLX$ + " LEFT JOIN MOVISTO M ON  M.DoSecLar = C.CodiCom_Lar "
    SQLX$ = SQLX$ + " where C.FechConta >= '" & DESDA & "' "
    SQLX$ = SQLX$ + " AND C.FechConta <= '" & HASDA & "' "
    SQLX$ = SQLX$ + " AND C.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " AND C.TipoMovim = 'VENTA' "
    cantidadRegistros& = CantRegistros("(" & SQLX$ & ") T ", "")
    SQLX$ = SQLX$ + " ORDER BY C.NumeComp ,C.FechEmisi ASC "
   


    Screen.MousePointer = 11
'    SQLX$ = SQLX$ + " ORDER BY CAJABANC.FechEmisi ASC, CAJABANC.NumeComp ASC "

    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = READSET(SQLX$)
    With AVANCETMP
        i& = 0
        Do While Not .EOF
            '
            'MUESTRA LA MAYORIA DE LOS IMPOÑRTES EN DOLARES
            Call TRACE(24, i&, cantidadRegistros&)
            '
            If XVALO(!NUMEORDPED) < 1 Then GoTo 49 'PARA QUE NO TRAIGA EL DESPIECE
            
            TipoComprobante$ = SAFETEXT$(!OpciMovim)
            Cliente& = XVALO(!NUCLIEN)
            razonSocial$ = SAFETEXT$(!raso_cli)
            CotizacionDolar# = XVALO(!COTIZA_DOLAR): If CotizacionDolar# < 1 Then CotizacionDolar# = 1
            IMPORTE# = XVALO(!ValAbsComp)
                  IMPORTEDOLARES# = XVALO(!ValAbsComp) / CotizacionDolar#
            If !ModoOpciMovim = "COMPENSA" Then IMPORTE# = 0 ' ES UN DEBITO POR COMPENSACION CLIENTE-PROVEEDOR
              If TipoComprobante$ = "NC" Then IMPORTE# = (-1) * IMPORTE#
            Comprobante$ = SAFETEXT$(!codicom_lar)
            
            FechaEmision% = CVINT(!fechconta)
            LOTE$ = SAFETEXT$(!idenlote)
            If XVALO(!MoneVal) >= 2 Then 'SI ESTA EN DOLARES MANTIENE
               PrecioUnitMov# = XVALO(!VALOUNIT)
            Else 'SI NO ESTA EN DOLARES CONVIERTE
               PrecioUnitMov# = XVALO(!VALOUNIT) / CotizacionDolar# 'SI ESTA EN PESOS LO PASA A DOLARES
            End If
            VALORFOB# = 0: GASTOSNACIONALIZACION# = 0
            If TRIM$(LOTE$) <> "" Then
                Call COSTOFOBYNACIONALIZADO(TRIM$(LOTE$), VALORFOB#, GASTOSNACIONALIZACION#) 'ESTA RUTINA LE DEVUELVE EL VALOR FOB Y EL VALOR FOB CON EL GASTO DE NACION. INCLUIDO
            End If
            SUVALORFOB# = VALORFOB#
            SUGASTOSNACIONALIZACION# = GASTOSNACIONALIZACION#
            'COSTO_NACIONALIZADO# = SUVALORFOB# + SUGASTOSNACIONALIZACION#
            COSTO_NACIONALIZADO# = GASTOSNACIONALIZACION#
            If COSTO_NACIONALIZADO# < 0.05 Then COSTO_NACIONALIZADO# = 1 'PARA EVITAR DIVISION POR 0
            DIF_PORCENTUAL = FORMATNUM$((PrecioUnitMov# - COSTO_NACIONALIZADO#) / COSTO_NACIONALIZADO# * 100, "F10.2") ' PORCENTAJE DIFERENCIA
            '
            i& = i& + 1
            GRILLAGRAL.GRID.Rows = i& + 1
            GRILLAGRAL.GRID.TextMatrix(i&, 1) = FECHATEX$(FechaEmision%)
            GRILLAGRAL.GRID.TextMatrix(i&, 2) = Comprobante$
            GRILLAGRAL.GRID.TextMatrix(i&, 3) = Cliente&
            GRILLAGRAL.GRID.TextMatrix(i&, 4) = SAFETEXT$(!raso_cli)
            GRILLAGRAL.GRID.TextMatrix(i&, 5) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(i&, 6) = SAFETEXT$(!idenlote)
            CANTI = XVALO(!CANTSALID)
            If CANTI <= 0 Then CANTI = XVALO(!CANTINGRE)
            GRILLAGRAL.GRID.TextMatrix(i&, 7) = CANTI
            
            GRILLAGRAL.GRID.TextMatrix(i&, 8) = FORMATNUM$(IMPORTE#, "F12.2") 'TOT FACTURA $ DETALLE
            GRILLAGRAL.GRID.TextMatrix(i&, 9) = FORMATNUM$(IMPORTEDOLARES#, "F12.2") 'TOT FACTURA DOLAR DETALLE
            
            GRILLAGRAL.GRID.TextMatrix(i&, 10) = FORMATNUM$(PrecioUnitMov#, "F12.2") 'PRECIO UNITARIO U$s DETALLE
            GRILLAGRAL.GRID.TextMatrix(i&, 11) = FORMATNUM$(PrecioUnitMov# * CANTI, "F12.2") 'IMPORTE ITEMS U$S(*1) DETALLE
            
            GRILLAGRAL.GRID.TextMatrix(i&, 12) = FORMATNUM$(SUVALORFOB# * CANTI, "F12.2") 'IMPORTE FOB DOLAR DETALLE
            GRILLAGRAL.GRID.TextMatrix(i&, 13) = FORMATNUM$(COSTO_NACIONALIZADO# * CANTI, "F12.2") 'COSTO NACIONALIZADO DETALLE
            
            GRILLAGRAL.GRID.TextMatrix(i&, 14) = FORMATNUM$(XVALO(DIF_PORCENTUAL), "F12.2") '% DETALLE 13,11
            GRILLAGRAL.GRID.TextMatrix(i&, 15) = FORMATNUM$(CotizacionDolar#, "F12.2")
            'nuevas columnas solicitadas por Fernando
            ResultadoBruto# = (PrecioUnitMov# * CANTI) - (COSTO_NACIONALIZADO# * CANTI): If ResultadoBruto# <= 0 Then ResultadoBruto# = 0
            GRILLAGRAL.GRID.TextMatrix(i&, 16) = FORMATNUM$(ResultadoBruto#, "F12.2")
            GRILLAGRAL.GRID.TextMatrix(i&, 17) = FORMATNUM$(ResultadoBruto# / (PrecioUnitMov# * CANTI) * 100, "F12.2")

    'CONSULTA SIEMPRE POR DETALLE SI SE ELIGE POR DETALLE MUESTRA EL RESULTADO DE ESTA CONSULTA
    'SI NO MUESTRA EL DICCIONARIO QUE ARMA DE ACUERDO A ESTA CONSULTA

            
            If Not diccionario.Exists(Comprobante$) Then
            
            ' Si no existe, agregar un nuevo registro
                diccionario.Add Comprobante$, Array( _
                FECHATEX$(FechaEmision%), _
                Comprobante$, _
                Cliente&, _
                razonSocial$, _
                , _
                , _
                CANTI, _
                IMPORTE#, _
                IMPORTEDOLARES#, _
                PrecioUnitMov#, _
                PrecioUnitMov# * CANTI, _
                FORMATNUM$(SUVALORFOB# * CANTI, "F12.2"), _
                FORMATNUM$(COSTO_NACIONALIZADO# * CANTI, "F12.2"), _
                FORMATNUM$(XVALO(DIF_PORCENTUAL), "F12.2"), _
                FORMATNUM$(CotizacionDolar#, "F12.2"), _
                FORMATNUM$(ResultadoBruto#, "F12.2"), _
                FORMATNUM$(ResultadoBruto# / (PrecioUnitMov# * CANTI) * 100, "F12.2"))
                 
            Else
                'MANERA PARA QUE ACUMULE LOS VALORES, HAY QUE RESTAR UNO A DIFERENCIA DE LA DE ARRIBA CUANDO ENTRA POR PRIMERA VEZ
                ' Si existe, acumular los valores
                 VALORES = diccionario(Comprobante$)
                 VALORES(10 - 1) = XVALO(VALORES(10 - 1)) + PrecioUnitMov#
                 VALORES(11 - 1) = XVALO(VALORES(11 - 1)) + (PrecioUnitMov# * CANTI)
                 VALORES(12 - 1) = XVALO(VALORES(12 - 1)) + (SUVALORFOB# * CANTI)
                 VALORES(13 - 1) = XVALO(VALORES(13 - 1)) + (COSTO_NACIONALIZADO# * CANTI)
                 
                 VALORES(16 - 1) = XVALO(VALORES(16 - 1)) + (ResultadoBruto#)
                 VALORES(17 - 1) = XVALO(VALORES(17 - 1)) + (ResultadoBruto# / (PrecioUnitMov# * CANTI) * 100)
                 'DE ESTA MANERA NO ACUMULA,
'                diccionario(COMPROBANTE$)(8) = IMPORTE# + XVALO(diccionario(COMPROBANTE$)(8)) ' Acumular importe
'                diccionario(COMPROBANTE$)(9) = XVALO(diccionario(COMPROBANTE$)(9)) + IMPORTEDOLARES#  ' Acumular importe
'                diccionario(COMPROBANTE$)(10) = XVALO(diccionario(COMPROBANTE$)(10)) + PrecioUnitMov# ' Acumular importe
'                diccionario(COMPROBANTE$)(11) = XVALO(diccionario(COMPROBANTE$)(11)) + (PrecioUnitMov# * CANTI) ' Acumular importe
'                diccionario(COMPROBANTE$)(12) = XVALO(diccionario(COMPROBANTE$)(12)) + (SUVALORFOB# * CANTI) ' Acumular importe
'                diccionario(COMPROBANTE$)(13) = XVALO(diccionario(COMPROBANTE$)(13)) + (COSTO_NACIONALIZADO# * CANTI) ' Acumular importe
                 diccionario(Comprobante$) = VALORES
            End If            '
49      .MoveNext
        Loop
      End With
      AVANCETMP.Close
      Set AVANCETMP = Nothing
      JJ& = i& 'PARA CONTTROLAR SI DEVOLVIO INFORMACION EL PERIODO SELECCIONADO
      
      If OPX% = 1 Then
        Screen.MousePointer = 1
        GRILLAGRAL.Show 1
        GoTo 99
      End If

      '
70    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)

      i& = 1
      GRILLAGRAL.GRID.Rows = diccionario.Count + 1 ' Establecer el número de filas en el grid
      GRILLAGRAL.GRID.Cols = 18
      Dim CLAVE As Variant
      For Each CLAVE In diccionario.Keys
        GRILLAGRAL.GRID.TextMatrix(i&, 0) = CLAVE ' Comprobante
        GRILLAGRAL.GRID.TextMatrix(i&, 1) = diccionario(CLAVE)(0)
        GRILLAGRAL.GRID.TextMatrix(i&, 2) = diccionario(CLAVE)(1)
        GRILLAGRAL.GRID.TextMatrix(i&, 3) = FORMATNUM$(diccionario(CLAVE)(2), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 4) = SAFETEXT$(diccionario(CLAVE)(3))
        GRILLAGRAL.GRID.TextMatrix(i&, 5) = SAFETEXT$(diccionario(CLAVE)(4))
        GRILLAGRAL.GRID.TextMatrix(i&, 6) = FORMATNUM$(diccionario(CLAVE)(5), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 7) = diccionario(CLAVE)(6)
        GRILLAGRAL.GRID.TextMatrix(i&, 8) = FORMATNUM$(diccionario(CLAVE)(7), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 9) = FORMATNUM$(diccionario(CLAVE)(8), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 10) = FORMATNUM$(SAFETEXT$(diccionario(CLAVE)(9)), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 11) = FORMATNUM$(diccionario(CLAVE)(10), "F12.2")

        GRILLAGRAL.GRID.TextMatrix(i&, 12) = FORMATNUM$(diccionario(CLAVE)(11), "F12.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 13) = FORMATNUM$(diccionario(CLAVE)(12), "F12.2")
        If OPX% = 1 Then  'DETALLE COMPARA POR EL IMPORTE DEL ITEM CONTRA EL IMPORTE FOB DEL ITEM
          DIF_PORCENTUAL_TOTAL = (XVALO(diccionario(CLAVE)(10)) - XVALO(diccionario(CLAVE)(12))) / XVALO(diccionario(CLAVE)(12)) * 100
        Else 'ACUMULADO:  COMPARA POR EL TOTAL DE LA FACTURA CONTRA EL IMPORTE TOTAL FOB
          DIF_PORCENTUAL_TOTAL = (XVALO(diccionario(CLAVE)(8)) - XVALO(diccionario(CLAVE)(12))) / XVALO(diccionario(CLAVE)(12)) * 100
        End If
        GRILLAGRAL.GRID.TextMatrix(i&, 14) = FORMATNUM$(DIF_PORCENTUAL_TOTAL, "F10.2")
        GRILLAGRAL.GRID.TextMatrix(i&, 15) = FORMATNUM$(diccionario(CLAVE)(14), "F12.2")
        If OPX% = 1 Then
            GRILLAGRAL.GRID.TextMatrix(i&, 16) = FORMATNUM$(diccionario(CLAVE)(15), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(i&, 17) = FORMATNUM$(diccionario(CLAVE)(16), "F12.2")
        Else
            GRILLAGRAL.GRID.TextMatrix(i&, 16) = FORMATNUM$(diccionario(CLAVE)(15), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(i&, 17) = FORMATNUM$(XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 16)) / XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 9)) * 100, "F12.2")
        End If
        i& = i& + 1
      Next CLAVE
      Screen.MousePointer = 1
      Erase RESP_ZELE_VECT
      GRILLAGRAL.Show 1
      
      NUDOCU$ = TRIM$(RESP_ZELE_VECT(2))
      NCLIEN& = XVALO(RESP_ZELE_VECT(3))
      
      If NUDOCU$ = "" Then GoTo 99
      Screen.MousePointer = 11
      NUDOCU1$ = TRIM$(Mid$(NUDOCU$, 9, 8))
      NUDOCU2$ = Mid$(NUDOCU$, 4, 13)
      TIPOCOM% = 1

      DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, VARCONTROL$)
      If DOCUNU& < 1 Then
        If TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            CONSIDERAR_MOVIVIE% = 1
            Call MUESTRASTSCREEN(NCLIEN&, NUDOCU$)
            CONSIDERAR_MOVIVIE% = 0
            Screen.MousePointer = 1
            GoTo 70
        Else
            Call COMUNI("Documento No Encontrado")
        End If
      Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
        Screen.MousePointer = 1
        GoTo 70
      End If
      Screen.MousePointer = 1
      
      If JJ& = 0 Then
        Call MENSERR(24, "No Se Ha Encontrado Información para el Período Elegido.")
        Exit Sub
      End If
      
99    Set diccionario = Nothing
      Screen.MousePointer = 1


End Sub



Private Sub mnukardex_Click()
  FIKARDEX09.Show 1
End Sub

Private Sub soloxproveedor_Click()

    CODI$ = TRIM$(Text1)
    CIII% = CODINT%(CODI$)
    '
    If CIII% < 1 Then
       Call COMUNI("Código Inexistentes o No Válido")
       Exit Sub
    End If
    '
    Call SELECHO("PROVED1")
    PROVED = XVALO(VALACT1$("PROVED1"))
    If PROVED > 0 Then
       Call CARUBICSQL(CIII%, PROVED)
    End If

End Sub
'
Private Sub stockdetallexproveedor_Click()

    CODXX$ = Text1.TEXT
    CODII% = CODINT%(CODXX$)
    If CODII < 1 Then
      Call COMUNI("Código Inexistente o No Válido")
      Exit Sub
    End If
    '
    Campos$ = "Cta.Prov/F9;R.Social/A40;O.Stock/F10.1;Ubicación/A48"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    '
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    FRMZELECHO.MSF1.Rows = 1
    J& = 0
    ACUMLA_SALDO = 0
    SQLX$ = " SELECT B.COD_EXTE, B.NUMPROV, M.COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO M with(nolock)"
    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP B ON M.IDENLOTE = B.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE B.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " AND M.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY M.Cod_Inte, B.NumProv,B.COD_EXTE"
    'SQLX$ = SQLX$ + " ORDER BY M.FECHMOV"
    Set rs1 = READSET(SQLX$)
    With rs1
        Do While Not .EOF
             NPROV = XVALO(!NumProv)
             RAZONPROVEEDOR$ = rasocli$(-1 * NPROV)
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             '
             J& = J& + 1
             FRMZELECHO.msf2.Rows = J& + 1
             FRMZELECHO.msf2.TextMatrix(J&, 1) = NPROV
             FRMZELECHO.msf2.TextMatrix(J&, 2) = RAZONPROVEEDOR$
             FRMZELECHO.msf2.TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             If NPROV > 0 Then
                FRMZELECHO.msf2.TextMatrix(J&, 4) = UBISTOSQL$(CODII%, NPROV)
             Else
                FRMZELECHO.msf2.TextMatrix(J&, 4) = UBISTO$(CODII%)
             End If
             
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    '
    'AGREGO AL FINAL EL ACUMULADO DE LOS MOVIMIENTOS DEL ITEM QUE SE GENERARON SIN LOTE POR EJEMPLO FACTURAS GENERADAS SIN LOTE.
    SQLX$ = " SELECT COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO With(nolock)"
    SQLX$ = SQLX$ + " WHERE (IDENLOTE IS NULL OR IDENLOTE = '') "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte"
    Set rs1 = READSET(SQLX$)
    With rs1
        Do While Not .EOF
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             '
             J& = J& + 1
             FRMZELECHO.msf2.Rows = J& + 1
             FRMZELECHO.msf2.TextMatrix(J&, 1) = "0"
             FRMZELECHO.msf2.TextMatrix(J&, 2) = "No Identificado"
             FRMZELECHO.msf2.TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             '
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Caption = "Código: " & CODXX$ & "     -     Saldo: " & ACUMLA_SALDO
    FRMZELECHO.Show 1
    
    PROVED = XVALO(RESP_ZELE_VECT(1))
    If PROVED > 0 Then Call VERUBICA(CODII%, PROVED)
    

End Sub


Sub VERUBICA(CIII%, PROVED)
    Call CARUBICSQL(CIII%, PROVED)
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

Sub ESTAXART(Optional CI1%)
   If CI1% < 1 Then
     Call SELECHO("MASTER")
     COD$ = TRIM$(VALACT1$("MASTER"))
     CI1% = CODINT%(COD$)
   End If
   
   If CI1% < 1 Then Exit Sub
   '
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'LISTA'"
   CONDI$ = CONDI$ + "  AND CODINT = " & CI1%
   CONDI2$ = "  AND MONEDA <= 1"
   CONDI3$ = "  AND MONEDA > 1"
   If CantRegistros("NOVEDADES", CONDI$ + CONDI2$, "NOMESS") > 0 Then Peso% = 1
   If CantRegistros("NOVEDADES", CONDI$ + CONDI3$, "NOMESS") > 0 Then DOL% = 1
   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
   If Peso% + DOL% > 1 Then
60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
      If OPX% < 1 Or OPX% > 2 Then GoTo 20
      TIPOMONE% = OPX%
   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
   Else
      TIPOMONE% = 1
      If DOL% > 0 Then TIPOMONE% = 2
   End If
   SQLX$ = "SELECT  * FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " ORDER BY FECH_MOD"
   '
   INFO_GRILLA07.MSF.Redraw = False
   Screen.MousePointer = 11
   Call CARGA_ENCABEZADO_X_ART
   REG& = 0
   Set rs = READSET(SQLX$)
   With rs
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLA07.MSF.Rows = REG& + 1
      Call TRACE(24, REG&, FIN&)
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!Fech_Mod))
      P_ANT# = XVALO(!Prec_Ant)
      P_NEW# = XVALO(!Prec_New)
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = FORMATNUM$(P_ANT#, "F12.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = FORMATNUM$(P_NEW#, "F12.2")
       PJE = ((P_NEW# / P_ANT#) - 1) * 100
      INFO_GRILLA07.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PJE, "F6.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = FORMATNUM$(TIPOMONE%, "F2.0")
       
      INFO_GRILLA07.MSF.TextMatrix(REG&, 6) = TRIM$(ECOARCH$("FLEXUSR", MKI$(XVALO(!Nume_Usuar))))
      INFO_GRILLA07.MSF.TextMatrix(REG&, 7) = SAFETEXT$(!OBSERVACION)

      .MoveNext
    Loop
   End With
   
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If
   
   INFO_GRILLA07.MSF.Redraw = True
   
   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)
   INFO_GRILLA07.Caption = "Cambios de Precios en " & DEMONE$ & "  - Art.: " & CODEXT$(CI1%) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HEFF1F1)
   '
   INFO_GRILLA07.Command7.Visible = True
   INFO_GRILLA07.CI1X = CI1%
   '
   INFO_GRILLA07.Show 1
   '
   If Peso% + DOL% > 1 Then GoTo 60
   '
   GoTo 20
   
End Sub

Sub ESTAXPRV()
   
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% < 1 Then Exit Sub
   '
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then GoTo 10

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'COSTO'"
   CONDI$ = CONDI$ + "  AND NumProv  = " & NCX1%
'   CONDI2$ = "  AND MONEDA <= 1"
'   condi3$ = "  AND MONEDA > 1"
   Peso% = 0: DOL% = 0
'   If CantRegistros("NOVEDADES", CONDI$ + CONDI2$, "NOMESS") > 0 Then PESO% = 1
'   If CantRegistros("NOVEDADES", CONDI$ + condi3$, "NOMESS") > 0 Then DOL% = 1
'   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
'   If PESO% + DOL% > 1 Then
'60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
'      If OPX% < 1 Or OPX% > 2 Then GoTo 20
'      TIPOMONE% = OPX%
'   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
'   Else
'      TIPOMONE% = 1
'      If DOL% > 0 Then TIPOMONE% = 2
'   End If
'   '
'   REG& = 0
   Call CARGA_ENCA_X_PROV
   
   '
   SQLX$ = "SELECT CODINT  FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
'   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
'   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " GROUP BY CODINT "
   '
   INFO_GRILLAJM.MSF.Rows = 1
   REG& = 0
   INFO_GRILLAJM.MSF.Redraw = False
   Screen.MousePointer = 11
   Set rs = READSET(SQLX$)
   With rs
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLAJM.MSF.Rows = REG& + 1
      '
      'LEO LA INFORMACION EN PARTICULAR POR CODIGO TENIENDO EN CUENTA ORDENAMIENTOS
      CI1% = XVALO(!CODINT)
      
      'TOMO LA FECHA PRIMERA DENTRO DEL RANGO
      Condi1$ = "CODINT = " & CI1%
      Condi1$ = Condi1$ + " and Fech_Mod  >=  '" & DESFE & " '"
      Condi1$ = Condi1$ + " and Fech_Mod  <=  '" & HASFE & " '"
      Condi1$ = Condi1$ + "  AND NumProv  = " & NCX1%
      Condi1$ = Condi1$ + " ORDER BY Fech_Mod ASC"
      FE1% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", Condi1$, "NOMESS"))
      If FE1% < 1 Then
        Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
        GoTo 20
      End If
      '
      'LEO EL  PRIMER PRECIO
      CONDI2$ = "CODINT = " & CI1%
      CONDI2$ = CONDI2$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE1%) & " '"
      CONDI2$ = CONDI2$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE1%, "HASTA") & " '"
      CONDI2$ = CONDI2$ + "  AND NumProv  = " & NCX1%
      CONDI2$ = CONDI2$ + " ORDER BY Fech_Mod ASC"
      PRECIO1# = XVALO(VALOBADA("NOVEDADES", "Prec_Ant", CONDI2$, "NOMESS"))
      '
      'LEO LA ULTIAMA FECHA DENTRO DEL RANGO
      CONDI3$ = "CODINT = " & CI1%
      CONDI3$ = CONDI3$ + " and Fech_Mod  >=  '" & DESFE & " '"
      CONDI3$ = CONDI3$ + " and Fech_Mod  <=  '" & HASFE & " '"
      CONDI3$ = CONDI3$ + "  AND NumProv  = " & NCX1%
      CONDI3$ = CONDI3$ + " ORDER BY Fech_Mod DESC"
      FE2% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", CONDI3$, "NOMESS"))
      '
      'LEO EL ULTIMO  PRECIO
      CONDI4$ = "CODINT = " & CI1%
      CONDI4$ = CONDI4$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE2%) & " '"
      CONDI4$ = CONDI4$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE2%, "HASTA") & " '"
      CONDI4$ = CONDI4$ + "  AND NumProv  = " & NCX1%
      CONDI4$ = CONDI4$ + " ORDER BY Fech_Mod DESC"

      PRECIO2# = XVALO(VALOBADA("NOVEDADES", "Prec_New", CONDI4$, "NOMESS"))
      '
      'CALCULO EL PORCENTAJE DE MODIFICACION ENTRE EL PRIMER VALOR Y EL ULTIMO SEGUN RANGO
      If PRECIO1# > 0.005 Then
         PJE = ((PRECIO2# / PRECIO1#) - 1) * 100
      ElseIf PRECIO2# < 0.005 Then
         PJE = -100
      'CASO QUE SEA INGRESO DE UN ITEM QUE SE CARGA EL PRECIO POR PRIMERA VEZ  X Q' CUANDO SE GRABE EN LA TABLA NOVEDADES NO VA A TRAER NINGUN COSTO ANTERIOR
      ElseIf PRECIO1# < 0.005 Then
         PJE = 0
      End If
      '
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 1) = CODEXT$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 2) = DESCRIT0$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 3) = FECHATEX$(FE1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PRECIO1#, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 5) = FECHATEX$(FE2%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 6) = FORMATNUM$(PRECIO2#, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 7) = FORMATNUM$(PJE, "F9.2")
'      INFO_GRILLAJM.MSF.TextMatrix(REG&, 8) = FORMATNUM$(TIPOMONE%, "F2.0")
      '
      .MoveNext
    Loop
   End With
   
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1

   INFO_GRILLAJM.MSF.Redraw = True
   
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If

   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)

   INFO_GRILLAJM.Caption = "Cambios de Precios en " & DEMONE$ & " - Prov.: " & rasocli(NCX1% * -1) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLAJM.MSF, &HC0E0FF)
   INFO_GRILLAJM.NPROV% = NCX1%
   INFO_GRILLAJM.Command7.Visible = True
   INFO_GRILLAJM.Show 1
   
   
   GoTo 20
End Sub




Sub CARGA_ENCA_X_PROV()
     With INFO_GRILLAJM
        .MSF.Rows = 1
        .MSF.Cols = 9
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        NUCOLU% = 2: .TEPRUEBA = "12345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Còdigo"
        NUCOLU% = 3: .TEPRUEBA = "123456789012345678901234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Descripciòn"
        NUCOLU% = 4: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 1"
        NUCOLU% = 5: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 6: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 2"
        NUCOLU% = 7: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 8: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        '
        NUCOLU% = 9: .TEPRUEBA = "12345": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        '
    End With
End Sub


'
Sub PRINTINFOGRILLAJM(NPROV%, Des%, Has%)
  
    If INFO_GRILLAJM.MSF.Rows < 1 Then
        Call COMUNI("Imposible Imprimir  - Sin Información")
        Exit Sub
    End If
    '
    Call BORRAREGISTROS("PRINTNOVEDADES", "CODINT = '' OR CODINT IS NULL OR CODINT > 0 ", REGAF&, "NOMESS")
    
    SQLX$ = "SELECT * FROM PRINTNOVEDADES "
    SQLX$ = SQLX$ + " WHERE CODINT = -1"
    '
    Set Rst = New ADODB.Recordset
25  On Error Resume Next
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0

  With INFO_GRILLAJM
          For i& = 1 To .MSF.Rows - 1
             CODI$ = TRIM$(.MSF.TextMatrix(i&, 1))
             FE1% = FECHANUM(.MSF.TextMatrix(i&, 3))
             P_ANT = XVALO(.MSF.TextMatrix(i&, 4))
             FE2% = FECHANUM(.MSF.TextMatrix(i&, 5))
             P_NEW = XVALO(.MSF.TextMatrix(i&, 6))
             PJE = XVALO(.MSF.TextMatrix(i&, 7))
             'MONEX% = XVALO(.MSF.TextMatrix(I&, 8))
             '
             With Rst
               .AddNew
               !CODINT = CODINT%(CODI$)
               !PORCENTAJE = PJE
                !Fech_Mod1 = CVDAT(FE1%)
                !Fech_Mod2 = CVDAT(FE2%)
                !Prec_Ant = P_ANT
                !Prec_New = P_NEW
                '!MONEDA = MONEX%
               .Update
             End With
   
          Next i&
  End With
  '
  FILTX$ = TRIM$(NPROV%) & ";" & TRIM$(Str$(Des%) & ";" & TRIM$(Str$(Has%)))
    
  Call STDFORM("NOVECOS", FILTX$)

  
        
End Sub


Sub PRINTINFOGRILLA07(CI1%, Des%, Has%)
  
    If INFO_GRILLA07.MSF.Rows < 1 Then
        Call COMUNI("Imposible Imprimir  - Sin Información")
        Exit Sub
    End If
    '
    Call BORRAREGISTROS("PRINTNOVEDADES", "CODINT = '' OR CODINT IS NULL OR CODINT > 0 OR CODINT <= 0", REGAF&, "NOMESS")
    '
    SQLX$ = "SELECT * FROM PRINTNOVEDADES "
    SQLX$ = SQLX$ + " WHERE CODINT = -1"
    '
    Set Rst = New ADODB.Recordset
25  On Error Resume Next
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With INFO_GRILLA07
          For i& = 1 To .MSF.Rows - 1
             FE1% = FECHANUM(.MSF.TextMatrix(i&, 1))
             P_ANT = XVALO(.MSF.TextMatrix(i&, 2))
             P_NEW = XVALO(.MSF.TextMatrix(i&, 3))
             PJE = XVALO(.MSF.TextMatrix(i&, 4))
             MONEX% = TRIM$(.MSF.TextMatrix(i&, 5))
             NOMBUSER$ = TRIM$(Left$(.MSF.TextMatrix(i&, 6), 24))
             OBSERV$ = Left$(.MSF.TextMatrix(i&, 7), 128)
             Imp_Mon_Orig = XVALO(.MSF.TextMatrix(i&, 8))
             COTIZA = XVALO(.MSF.TextMatrix(i&, 9))
             '
             With Rst
                .AddNew
                !PORCENTAJE = PJE
                !Fech_Mod1 = CVDAT(FE1%)
                !Prec_Ant = P_ANT
                !Prec_New = P_NEW
                !OBSERVACION = OBSERV$
                !Moneda = MONEX%
                !NOM_USER = NOMBUSER$
                !COTIZA_DOLAR = COTIZA
                !IMP_ORG_DOLAR = Imp_Mon_Orig
                
               .Update
             End With
   
          Next i&
    End With
    '
    FILTX$ = TRIM$(CI1%) & ";" & TRIM$(Str$(Des%) & ";" & TRIM$(Str$(Has%)))
    '
    Call STDFORM("NOVECOS2", FILTX$)

  
        
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
'        On Error Resume Next
'
'        MSF(0).Height = Me.Height - 1600
'        MSF(0).Width = Me.Width - 235
'        MSF(1).Height = Me.Height - 1600
'        MSF(1).Width = Me.Width - 235
'        MSF(2).Height = Me.Height - 16001
'        MSF(2).Width = Me.Width - 235
'        MSF(3).Height = Me.Height - 1600
'        MSF(3).Width = Me.Width - 235
'        MSF(4).Height = Me.Height - 1600
'        MSF(4).Width = Me.Width - 235
'
'        SSTab1.Height = Me.Height - 1600
'        SSTab1.Width = Me.Width - 235
'        On Error GoTo 0
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
   If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
   End If
   
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
      archivo$ = "TABTIMB"
     Else
      archivo = "TATIMAT"
   End If
   '
   Call SELECHO(archivo$)
   VALONUE$ = VALACT1$(archivo$)
   If VALONUE$ <> "" Then
      Text6(0).TEXT = VALONUE$
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
    Label41.Caption = Unimed$(CDYY%)
    
    
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

Private Sub Command21_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   If DERACCE%("ALTAITEM", "Alta de un Nuevo Item de Stock") < 2 Then
     Exit Sub
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
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
       CI_NEW = CODINT_NEW
       'Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
       Call REPLA(VDEF$, MKI$(CI_NEW), 1, 2)
       Call REPLA(VDEF$, "U.", 109, 2)
       Call REPLA(VDEF$, Chr$(1), 101, 1)
       '
       OKX% = 0
       Call ALTACODIGO(VDEF$, OKX%)
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
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
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   Dim FIELDXVALO(256)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
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
   CINUE% = CODINT_NEW
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, MKI$(GRUCOVEI%(CII%)), 3, 2)
   Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
   Call REPLA(VDEF$, Unimed$(CII%), 109, 2)
   Call REPLA(VDEF$, Chr$(TIMATE%(CII%)), 101, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
     Dim RMaster As ADODB.Recordset
     Set RMaster = New ADODB.Recordset
     strSQL$ = "Select * from Master WITH(NOLOCK) where codint=" & CII%
     Set RMaster = READSET(strSQL$)
     If Not (RMaster.EOF And RMaster.BOF) Then
        CACAM% = RMaster.Fields.Count
        For KU% = 1 To CACAM%
          Dato$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
          NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
          If NOCAMPO$ <> "CODIGO" Then
             If NOCAMPO$ <> "DESCRIPCION" Then
                If NOCAMPO$ <> "CODINT" Then
                   If NOCAMPO$ <> "ALIAS" Then
                      If NOCAMPO$ <> "UMEDIDA" Then
                        CONDI$ = "CODINT=" & CINUE%
                        Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, Dato$, REGAF&, "NOMESS")
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
   If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
   End If
   
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
   
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% > 0 Then
     SQLX$ = "Update Master SET Codigo = '" & CODINUE$ & "' WHERE Codint = " & Abs(Ci%)
     FLEXCONN.Execute SQLX$
     UPDATMASTER% = 1
     Call GENECOMAS
     RECODI$ = CODEXT$(Ci%)
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "Codint = " & Abs(Ci%)
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Master", "MINEPARTS", "Master", CONDI$)
      End If
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
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   Command24.Enabled = False
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
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
      Set rs = READSET(SQLX$)
      II& = 0
      Do While Not rs.EOF
            Codigo$ = CODEXT$(Abs(XVALO(rs!CODICONJ)))
            LISTACODIGOS$ = LISTACODIGOS$ & Codigo$ & ";"
            II& = II& + 1
            If II& = 4 Then
                LISTACODIGOS$ = LISTACODIGOS$ & "\"
                II& = 0
            End If
        rs.MoveNext
      Loop
      'SE CAMBIA MENSAJE PORQUE SE HA DEJADO LA POSIBILIDAD DE ELIMINAR EL ITEM
      Call MENSERR(24, "El Codigo que Intenta dar de Baja\Es Componente de la Fórmula de Producto (Activo o de Baja)\De los Productos Padre Correspondientes:\ \" & LISTACODIGOS$)
      'GoTo 99
   End If
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(Ci%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   CII% = CODINT%(Text1.TEXT)
   If CII% > 0 Then
     strSQL = "Update Master set Status=-1 where Codint=" & Ci%
     FLEXCONN.Execute strSQL, cdadreg
     UPDATMASTER% = 1
     Call BLANFORMU
     Text1 = ""
     Call GENEECOMABAJ 'ACTUALIZA LISTA DE SELECCION
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "Codint=" & Ci%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
     End If
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
   MAXCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   
   For IX% = 0 To List1.ListCount - 1
     Codigo$ = TRIM$(Mid(List1.List(IX%), 1, MAXCODI%))
     If Codigo$ = TRIM$(Text31) Then
        RTA1% = COMALTER%("Codigo " & Codigo$ & " Pre-Existente en Fórmula\¿Realmente Desea Agregarlo Repetido?\\No, Cancelar\Si, Repetir")
        If RTA1% <> 2 Then Exit Sub
        Exit For
     End If
   Next IX%
 
   '--Fin Validacion Item en formula------------*
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = Unimed$(CIY%)
   Uso = XVALO(Text32.TEXT)
   REFCO$ = TRIM$(Text39)
   If Abs(Uso) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(Uso), 4, 12, 4, 2)
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
   MAXCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   
   For IX% = 0 To List1.ListCount - 1
     Codigo$ = TRIM$(Mid(List1.List(IX%), 1, MAXCODI%))
     If Codigo$ = TRIM$(Text31) Then
        RTA1% = COMALTER%("Codigo " & Codigo$ & " Pre-Existente en Fórmula\¿Realmente Desea Agregarlo Repetido?\\No, Cancelar\Si, Repetir")
        If RTA1% <> 2 Then Exit Sub
        Exit For
     End If
   Next IX%
 
   '--Fin Validacion Item en formula------------*
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = Unimed$(CIY%)
   Uso = XVALO(Text32.TEXT)
   REFCO$ = TRIM$(Text39)
   If Abs(Uso) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(Uso), 4, 12, 4, 2)
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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then
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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then
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
   RESECO$(List3.ListCount) = Left$(MKI$(Ci%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(Ci%) + MKI$(NOPY%) + String$(128, 0) 'para acomodar hoja de procesos DE FABRICACION
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
   If SSTab1.Tab <> 4 Then
        PRINTFORM
   End If
'   UTILIZA_SKIN% = 1
99 Command39.Enabled = True

'    Dim Nuevo_Form As Form
'
'    Set Nuevo_Form = New AMASTOK07
'
'    UTILIZA_SKIN% = 0
'    With Nuevo_Form
'        .PRINTFORM
'        .Hide
'    End With
'    '
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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then
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
   RESECO$(ILISTAX&) = Left$(MKI$(Ci%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(Ci%) + MKI$(NOPY%) + String$(128, 0) 'para acomodar hoja de procesos DE FABRICACION
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
   Call ZELECHOAHP(1)
   COD$ = RESP_ZELE_VECT(2)

'   Call SELECHO("MASTER")
'   COD$ = TRIM$(VALACT1$("MASTER"))

   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   
'   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
'   Call GRADATITEM
'   '
'   FORSELEC.ABUSQUE = Text1
'   Call SELECHO("MASTER")
'   COD$ = TRIM$(VALACT1$("MASTER"))
'   If COD$ <> "" Then
'     Text1.TEXT = COD$
'   End If
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
       Call ShellExecute(Me.hwnd, "Open", TRIM$(Label63), "", "", 1)
       'Call MUESTRAPDF(TRIM$(Label63))
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
   Linea$ = List2.List(List2.ListIndex)
   Archi$ = TRIM$(Mid$(Linea$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If
End Sub



Private Sub mnuInic_novedades_Click()
   '
   'INICIALIZACION DE COSTOS EN TABLA NOVEDADES
   'RECORRER TODO EL MASTER Y VALIDAR SI EL CODIGO NO EXISTE EN TABLA NOVEDADES AGREGARLO
   Screen.MousePointer = 11
   '
   CONDI$ = "CODINT > 0 AND Status >= 0 AND COSUNI > 0"
   FIN& = CantRegistros("MASTER", CONDI$, "NOMESS")
   
   i& = 0
   SQLX$ = "SELECT * FROM MASTER WITH(NOLOCK) WHERE CODINT > 0 "
   SQLX$ = SQLX$ + "  AND Status > 0 AND COSUNI > 0.00005 "
   Set DOCTMP = READSET(SQLX$)
   With DOCTMP
      Do While Not .EOF
        i& = i& + 1
        Call TRACE(20, i&, FIN&)
        CI1% = XVALO(!CODINT)
        NPROV% = XVALO(!PROVHABI)
        PRECIOX# = XVALO(!COSUNI)
        PRQ = FORMATNUM$(PRECIOX, "F12.2")
        MON% = XVALO(!Monecos)
        CONDI$ = "CODINT = " & CI1%
        If CantRegistros("NOVEDADES", CONDI$, "NOMESS") < 1 Then
          Call ADD_NOVEDAD(CI1%, NPROV%, PRECIOX#, PRECIOX#, "Inicialización de Costo en Novedades ", MON%, "COSTO")
        End If
        .MoveNext
      Loop
   End With
   '
   DOCTMP.Close
   Set DOCTMP = Nothing
   Screen.MousePointer = 1
   
   Call COMUNI("Proceso de 'Incialización de Costos en Tabla Novedades' Finalizado con Exito!")
   
    

End Sub


Private Sub MSF_DblClick(Index As Integer)
   '
   
   If Index = 0 Then
      REG& = XVALO(MSF(0).MouseRow)
      If REG& < 1 Then Exit Sub
       'MUESTRA POR  PANTALLA LA IMAGEN DEL DOCUMENTO ELEGIDO
       '****************************************************
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       'TIPO DE COMPROBANTE
        PRIMERACOLUM$ = UCase$(MSF(0).TextMatrix(REG&, 1))
       TIP$ = UCase$(Mid$(MSF(0).TextMatrix(REG&, 1), 1, 2))
       
       If TRIM$(TIP$) = "" Then Exit Sub
       '
       NUDOCU$ = TRIM$(Mid$(PRIMERACOLUM$, 4, 6))
       Select Case TIP$
         Case "PC"
           'PEDIDO ' TENER EN CUENTA DERECHOS PARA VISUALIZAR PRECIOS SI NO PRESENTAR ORDEN DE DESPACHO
10          OPX% = ALTERNA%("Visualizar Pedidos - Orden de Venta\Visualizar Orden de Despacho")
          If OPX% < 1 Or OPX% > 2 Then Exit Sub
          '
          If OPX% <> 2 Then
            If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
               Call MENSERR(24, "Usuario No Autorizado para Visualizar Pedidos.\Seleccione Orden de Despacho")
               GoTo 10
            End If
          End If
          '
          If OPX% = 1 Then
             TIDOCUM$ = "Pedido de Cliente"
             DOCUMENTO_ELEGIDO$ = "Orden de Venta"
          ElseIf OPX% = 2 Then
             TIDOCUM$ = "Orden de Despacho"
             DOCUMENTO_ELEGIDO$ = TIDOCUM$
          End If
          While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
          NUDOCU$ = NUDOCU$ + " "
          'FIN PEDIDO / ORDEN DE DESPACHO
          '
          Case "OF"
            'ORDEN DE FABRICACION
            While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
            NUDOCU$ = "OF-" + NUDOCU$ + "-" + "1"
            TIDOCUM$ = "Orden de Fabricación " + NUDOCU$
         Case "OC"
            'ORDEN DE COMPRA
            TIDOCUM$ = "Nota de Pedido"
            If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "ORDEN DE COMPRA"
            While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
            NUDOCU1$ = " " + NUDOCU$ + " "
            NUDOCU2$ = " " + NUDOCU$ + "-"
         Case "BR"
             'BOLETA DE RECEPCION
             TIDOCUM$ = "Boleta de Recepción"
             NUDOCU$ = " " + NUDOCU$ + " -"
         Case Else
             Exit Sub
       End Select
       '
        Screen.MousePointer = 11
        '
        For UI& = ULTREG&("PDOCLST") To 1 Step -1
            XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
            If InStr(XX$, TIDOCUM$) > 4 Then
              PPXX% = InStr(XX$, "Nro.")
              '
              If PPXX% > 0 Then
                XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
              End If
              '
              If TIP$ = "OC" Then
                'CASO DE ORDEN DE COMPRA BUSCA POR UN OR CON DOS VARIABLES
                If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
                  DOCUNU& = CVS(Left$(XX$, 4))
                  GoTo 5
                End If
              ElseIf InStr(XX$, NUDOCU$) > 4 Then
                 'CASOS RESTANTES
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 5
              End If
            End If
        Next UI&
        Screen.MousePointer = 1
        Call COMUNI("Documento no Encontrado")
        Exit Sub
        '
        'MUESTRA DOCUMENTO
        Screen.MousePointer = 1
5       Call MUESTRADOC(DOCUNU&)
   End If
   If Index = 1 Then
         REG& = XVALO(MSF(1).MouseRow)
         If REG& < 1 Then Exit Sub
         CUENTAPRO = XVALO(MSF(1).TextMatrix(REG&, 1))
         CI1% = XVALO(Label28)
         If CUENTAPRO > 0 Then
            Call CARUBICSQL(CI1%, CUENTAPRO)
            'REFRESCO
            Label15(2) = ""
            Label15(2) = Label28
         End If
   End If
   '
   If Index = 10 Then
     REG& = XVALO(MSF(10).MouseRow)
     If REG& < 1 Then Exit Sub
     
     CI1% = XVALO(Label28)
     ''' VERIFICAR SI HAY ITEMS RECIBIDOS PENDIENTES DE FACTURACION PARA QUITAR DE LA LISTA DE PENDIENTES EN FACTURACIOON DE PROVEEDORS
     'ESTA CONSULTA RETORNA LOS ITEMS QUE ESTAN EN BOLRECEP CON SU IDENLOTE , PERO QUE AUN NO ESTAN EN LA TABLA DETFACOMSQL
     SQLX$ = "SELECT COUNT(*) FROM BOLRECEP B WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT 1 FROM DETFACOMSQL D"
     SQLX$ = SQLX$ + " Where B.Cod_Inte = D.Cod_Inte And B.idenlote = D.idenlote)"
     Set RSCOUNT = READSET(SQLX$)
     If Not RSCOUNT.EOF Then
        cantidadRegistros = RSCOUNT.Fields(0).Value
     Else
        cantidadRegistros = 0 ' No hay registros
     End If
     If cantidadRegistros > 0 Then
        OPXX% = COMALTER%("Opción de Quitar Recepción de Pendientes de Factura\\Seleccionar y Quitar\Continuar con ingreso de precio de venta ")
        If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
        If OPXX% = 1 Then
        
             SQLX$ = "SELECT Cod_Inte, idenlote,FECHRECEP,CANTOAPRO FROM BOLRECEP B WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT 1 FROM DETFACOMSQL D"
             SQLX$ = SQLX$ + " Where B.Cod_Inte = D.Cod_Inte And B.idenlote = D.idenlote)"
             SQLX$ = SQLX$ + "  AND CANTOAPRO> 0 ORDER BY FECHRECEP DESC"
             Campos$ = "ci/F0;Fecha/D8;Código/A20;Descripción/A48;Lote/A12;Cant.Rec./F8"
             Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
             Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
             FRMZELECHO.Width = 12000
             FRMZELECHO.Caption = "Seleccione c/Doble Clic para Quitar de Lista de Pendientes de Facturación"
        
             Set rs = READSET(SQLX$)
             FRMZELECHO.msf2.Redraw = False
             With rs
              Do While Not .EOF
                i& = i& + 1
                FRMZELECHO.msf2.Rows = i& + 1
                FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!cod_inte)
                FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!FECHRECEP)
                FRMZELECHO.msf2.TextMatrix(i&, 3) = CODEXT$(XVALO(!cod_inte))
                FRMZELECHO.msf2.TextMatrix(i&, 4) = DESCRIT0$(XVALO(!cod_inte))
                FRMZELECHO.msf2.TextMatrix(i&, 5) = SAFETEXT$(!idenlote)
                FRMZELECHO.msf2.TextMatrix(i&, 6) = FORMATNUM$(XVALO(!CANTOAPRO), "F9.1")
                
                .MoveNext
              Loop
            End With
            rs.Close
            Set rs = Nothing
            FRMZELECHO.msf2.Redraw = True
            '
            FRMZELECHO.Show 1
            
            CI1% = XVALO(RESP_ZELE_VECT(1))
            If CI1% > 0 Then
               LOTE$ = RESP_ZELE_VECT(5)
               Call AgregarRegistroDetFacomSql(CI1%, LOTE)
               Call COMUNI("TRANSACCION REALIZADA")
            End If
            Exit Sub
        End If
        If OPXX% = 2 Then GoTo 50
   Else
        On Error Resume Next
50        PREVENTAX = TRIM$(MSF(10).TextMatrix(REG&, 7))
          If Err <> 0 Then On Error GoTo 0: Exit Sub 'SI SE PRODUCE ERROR DE INDICE FUERA DE TERMINO SE VA
          On Error GoTo 0 'SI NO CONTINUA
          
        PREVENTA_MODIF = TRIM$(InputBox$("Informar", "Ingrese Precio de Venta", PREVENTAX))
        If XVALO(PREVENTA_MODIF) > 0 Then
           MSF(10).TextMatrix(REG&, 7) = FORMATNUM$(XVALO(PREVENTA_MODIF), "F12.2")
           Call GRABALISTAAHP
        Else
           Call COMUNI("Imposible Modificar !!!\Valor de Venta Debe ser Mayor a 0")
        End If
     End If
    '       Call ACTUREGISTRO("COSTOXPROVE", "PRECIOVENTA", CONDI$, FORMATNUM$(XVALO(PREVENTA_MODIF), "F16.5"), REGAF&, "NOMESS")
   End If
   '
End Sub



Sub GRABALISTAAHP()
       CI1% = XVALO(Label28)
       '
       CONDI$ = "CODINT = " & CI1%
       FLEXCONN.BeginTrans
       SQLX$ = "DELETE FROM COSTOXPROVE WHERE " & CONDI$
       SQLX$ = SQLX$ + " AND CodiEmpr = " & CodiEmp%
       FLEXCONN.Execute (SQLX$)
       For i& = 1 To MSF(10).Rows - 1
          CTAROV = MSF(10).TextMatrix(i&, 1)
          prexiventa = XVALO(MSF(10).TextMatrix(i&, 7))
          If prexiventa > 0 Then prexiventa = FORMATNUM$(prexiventa, "F16.5")
          VIAENT% = MSF(10).TextMatrix(i&, 8)
          NLISTX% = 1
          
          SQLX$ = "INSERT INTO COSTOXPROVE "
          SQLX$ = SQLX$ + " (Nume_Cli,CODINT,LISTA,PRECIOVENTA,Via_Entrega,MARBORRA,CODIEMPR)"
          SQLX$ = SQLX$ + " VALUES(" & CTAROV
          SQLX$ = SQLX$ + "," & CI1%
          SQLX$ = SQLX$ + "," & NLISTX%
          SQLX$ = SQLX$ + "," & prexiventa
          SQLX$ = SQLX$ + "," & VIAENT%
          SQLX$ = SQLX$ + ", 0,0)"
          FLEXCONN.Execute (SQLX$)
       Next i&


ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Precio de Venta Guardado Correctamente")
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
  CI1% = CODINT%(CODEX1$)
  Call CALCUKILOS(CI1%)
  '
End Sub
Sub TEXT42_CHANGE()
  '
  CODEX1$ = Text31
  CI1% = CODINT%(CODEX1$)
  Call CALCUKILOS(CI1%)
  '
End Sub
Sub LABEL42_CHANGE()
  '
  CODEX1$ = Text31
  CI1% = CODINT%(CODEX1$)
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
   Ci% = CODINT%(Text1)
   If Ci% > 0 Then
     Call TRALOFIL("RECAMPH", 1, MKI$(Ci%))
     If ULTREG&("!RECAMPH") < 1 Then
          Call MENSERR(24, "Sin Cambios Registrados\para Código " + TRIM$(CODEXT$(Ci%)))
        Else
          VTB% = VENTABU%("RECAMPH/TITUPAN=Cambios Proveedor Habitual Código " + TRIM$(CODEXT$(Ci%)))
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
         CLAVE$ = MKI$(Ci%) + MKI$(NOPE%)
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
            X$ = REGLEIDO("!TACATMU", II&)
            COTMU$ = Mid$(X$, 1, 4)
            CANTI1 = CVD(Mid$(X$, 13, 8))
            MSEG1# = MSTMU#(COTMU$) 'ME DEVUELVE LOS MILISEGUNDOS DEL TMU
            TISTOP = TISTOP + (CANTI1 * MSEG1# / 1000)
            '
            With TMUSSSOPER
              .AddNew 'CARGO TODOS   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
                !cod_inte = CIXI%
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
      Label9 = Unimed$(CDYY%)
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
   Text6(0).TEXT = ""
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
   
   MSF(10).Rows = 1

End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    Ci% = CODINT%(CODD$)
    If Ci% < 1 Then
       If TRIM$(CODD$) <> "" Then
         If Ci% < 0 Then
             Call MENSERR(24, "Codigo Anulado o de Baja")
           Else
             Call MENSERR(24, "Código Inexistente o no Válido")
         End If
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text1.TEXT = CODEXT$(Ci%)
    Text3.TEXT = DESCRIT0$(Ci%)
    Text38.TEXT = DESCRIT2$(Ci%)
    Label37.Caption = Unimed$(Ci%)
    Label28.Caption = Str$(Ci%)
    Text45 = NUPLANO$(Ci%)
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
      Ci% = CODINT%(CODD$)
      If Ci% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
         UNID$ = Unimed$(Ci%)
      Text4.TEXT = UNID$
      Label30.Caption = ECOARCH$("UMEDIDA", UNID$)
         TIMA% = TIMATE%(Ci%)
      Text6(0).TEXT = TRIM$(Str$(TIMA%))
      Label29.Caption = ECOARCH$("TATIMAT", Chr$(TIMA%))
         FAMI$ = CODFAMIL$(Ci%)
      Text48.TEXT = FAMI$
      Label64.Caption = ECOARCH$("TAFAMIL", FAMI$)
         GRUCO% = GRUCOVEI%(Ci%)
      Text16.TEXT = TRIM$(Str$(GRUCO%))
      Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(GRUCO%))
         GRUIVA% = GRUPIVA%(Ci%)
      Text35.TEXT = TRIM$(Str$(GRUIVA%))
      Label57.Caption = ECOARCH$("GRUPIVA", Chr$(GRUIVA%))
      '
      ' SI TIENE DERECHO DE VER EL COSTO - SI NO PRESENTA ASTERIZCOS
      'PUEDE VER PERO LA CAJA DE TEXTO SE DESHABLITA
      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") > 0 Then
        Text12.TEXT = CSTRING$(MKS$(COSUNI(Ci%)), 3, 11, 4, 2)
        Text12.Enabled = False
        Command15.Enabled = False
        Text13.Enabled = False
        Text21.Enabled = False

        'SI TIENE DERECHOS DE MODIFICAR COSTOS SE HABILITA LA CAJA DE TEXTO
        If DERACCE%("MODCOSTO/NOMESS", "Modificar Costos y Precios") > 1 Then
           Text12.Enabled = True
           Command15.Enabled = True
           Text13.Enabled = True
           Text21.Enabled = True
        End If
      Else
        'PRESENTA ASTERIZCOS PARA QUE NO SE VISUALICE EL PRECIO
        Text12 = FORMATNUM$(99999999999#, "F11.4")
        Text12.Enabled = False
        Command15.Enabled = False
        Text13.Enabled = False
      End If

         FEX = FECOSTO%(Ci%)
      Text14.TEXT = FECHATEX$(FEX)
         MONEI% = MONECOSTO%(Ci%)
      Text13.TEXT = TRIM$(Str$(MONEI%))
      Label35.Caption = ECOARCH$("TAMONED", Chr$(MONEI%))
         PORNA = PORNACIO(Ci%)
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
         NPRO% = PROVHABI%(Ci%)
      Text10.TEXT = TRIM$(Str$(NPRO%))
      Label33.Caption = ECOARCH$("ACREDOR", MKI$(NPRO%))
         '
         CRIT% = CRITREP%(Ci%)
      Text8.TEXT = TRIM$(Str$(CRIT%))
      Label32.Caption = ECOARCH$("TACRIRE", Chr$(CRIT%))
      Text19.TEXT = CSTRING$(MKS$(LOTESTAN(Ci%)), 3, 10, 1, 2)
      Text18.TEXT = CSTRING$(MKS$(STOMIN(Ci%)), 3, 10, 1, 2)
      '
      '\\\OT-05-0720-WAR-25/11/05///
      On Error Resume Next
      Text40.TEXT = CSTRING$(MKS$(LINDIVIS(Ci%)), 3, 10, 1, 2)
      Text49.TEXT = CSTRING$(MKS$(LINDIVEN(Ci%)), 3, 10, 1, 2)
      On Error GoTo 0
      '\\\OT-05-0720-WAR-FIN///
      '
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(Ci%)), 3, 10, 1, 2)
      '
      CONDI$ = "Codint = " & Ci
      Text22.TEXT = COLOR$(Ci%)
      Text23.TEXT = TALLE$(Ci%)
      Text44.TEXT = VALOBADA("MASTER", "COD_INTERNA", CONDI$, "NOMESS")
      Text24.TEXT = LTITPE$(PESUNI(Ci%))
      '
      '\\\OT-5-0633-FG-21/09/05
      Text46.TEXT = FORMATNUM$(PESMETRO(Ci%), "F8.3")
      '\\\OT-5-0633-FG-21/09/05
      '
      Text36 = REVICOD$(Ci%)
      Text25 = FECHATEX$(FEREVI%(Ci%))
      '
      '\\\OT-06-0105-WAR-23/03/06///
      If DESCOMPO%(Ci%) = 1 Then Check3.Value = 1
      '\\\OT-06-0105-WAR-FIN///
      Label58 = FORMATNUM$(ESTRAZABLE%(Ci%), "I1")
      '
      Text50.TEXT = SAFETEXT(VALOBADA("MASTER", "CANDECI", CONDI$, "NOMESS"))
      CONS# = XVALO(VALOBADA("MASTER", "consuprom", CONDI$))
      Label67 = TRIM$(FORMATNUM(CONS#, "F13.1"))
      '
      COSCAL# = XVALO(VALOBADA("MASTER", "CosCal", CONDI$)) 'Costo Calculado
      Label69 = FORMATNUM(COSCAL#, "F13.4")
      '
      Label70 = FECHATEX(CVINT(VALOBADA("MASTER", "FECOSCAL", CONDI$)))
      '
      Label71 = "1"
      If STATMOVI%(Ci%) = -1 Then Label71 = ""
      If STATMOVI%(Ci%) = 1 Then Label71 = "2"
      '
      'Se modificaron estas dos lecturas para que traigan solo el dato de la base
      'Muestra el Valor del Lead_Time ingresado
      
      Text51.TEXT = VALOBADA("MASTER", "LeTiCom", CONDI$, "NOMESS")
      '
      'PERIODO DE STOCKS MINIMO
      'no se lee con xvalo para que no  traiga 0 en el caso que este vacio
      Text52.TEXT = VALOBADA("MASTER", "Autono_Proy", CONDI$, "NOMESS")
      '
      Call CARGACOSTOXPROVE

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
      Ci% = CODINT%(CODD$)
      If Ci% < 1 Then Exit Sub
      '
      Text43.TEXT = ESPECIFIT$(Ci%)
      Label63.Caption = ARMEMOTEC$(Ci%)
      '
      Text81 = TRIM$(VALOBADA("Master", "Clas_ATC", "codint = " & Ci%, "NOMESS"))
      Text82 = TRIM$(VALOBADA("Master", "Vers_BC", "codint = " & Ci%, "NOMESS"))
      Text83 = TRIM$(FECHATEX$(CVINT(VALOBADA("Master", "FeVig_BC", "codint = " & Ci%, "NOMESS"))))
      Text84 = TRIM$(FECHATEX$(CVINT(VALOBADA("Master", "FeRev_BC", "codint = " & Ci%, "NOMESS"))))
      YACAR%(1) = 1
      '
      List2.Clear
      SQLX$ = "SELECT * FROM DOCUREL WITH(NOLOCK) WHERE COD_INTE = " & Ci
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           RUTAX$ = SAFETEXT$(!Ruta$)
           Call REPLA(TX$, DESCRI$, 1, 24)
           Call REPLA(TX$, RUTAX$, 32, 512)
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
      Ci% = CODINT%(CODD$)
      If Ci% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      PORMILLAR% = 0
      'SE AGREGO PARAMETRO 'COMPLETO' A LEEEXPLO PARA QUE CUANDO ES 1 INCLUYE TAMBIEN A LOS ANULADOS
      Call LEEEXPLO(CODD$, 1, 1)
      Label59.Caption = FECHATEX$(FEREFOR%(Ci%))
      If TRIM$(Label59.Caption) <> "" Then
         CONDI$ = "CODINT = " & Ci
         Label59.ToolTipText = "Aprobada por " & SAFETEXT$(VALOBADA("Master", "UsuApFor", CONDI$, "NOMESS"))
      End If
      Label59.Refresh
      List1.Clear
      NOCHECK% = 1
      Check1.Value = 0: If PORMILLAR% = 1 Then Check1.Value = 1
      NOCHECK% = 0
'     Decide si muestra o no el Boton aprobado
      If Control("FORMULAS", "REQAPRO") = "SI" Then
         FRAME15(0).Caption = "Aprobacion"
         USU$ = SAFETEXT(VALOBADA("MASTER", "USUAPFOR", "CODINT = " & Ci%))
         If USU$ = "PENDIENTE DE APROBACION" Then
              Command51.Visible = True
            Else
              Command51.Visible = False
              Label59.Caption = FECHATEX$(FEREFOR%(Ci%))
              Label59.ToolTipText = "Aprobada por " & USU$
         End If
      End If
      
      '
      For i% = 1 To CAIT%
         List1.AddItem EXPLOLIN$(i%)
      Next i%
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
      Ci% = CODINT%(CODD$)
      If Ci% < 1 Then Exit Sub
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
      SQLX$ = SQLX$ & " WHERE CodIConj = " & Ci
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
          NOPY% = !NUMEOPER
          ALTY% = !AltOper
          DOPE$ = !descoper
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
          CATOP% = !CATEGOPS
          CATOP2% = !CATEGOPS2
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
          Label60 = FECHATEX$(CVINT(!Fech_Revis))
          Label62 = SAFETEXT$(!USUA_REVIS)
          '
          X$ = REBLA$
          Call REPLA(X$, MKI$(Ci%), 1, 2)
          Call REPLA(X$, MKI$(NOPY%), 3, 2)
          Call REPLA(X$, Chr$(ALTY%), 88, 1)
          Call REPLA(X$, DOPE$, 5, 48)
          Call REPLA(X$, EQUIPO$, 53, 6)
          Call REPLA(X$, MKS$(HORPREP), 59, 4)
          Call REPLA(X$, MKS$(HORFIJO), 63, 4)
          Call REPLA(X$, MKS$(MINVARI), 67, 4)
          Call REPLA(X$, MKS$(DEMODIS), 75, 4)
          Call REPLA(X$, MKS$(HORLIMP), 79, 4)
          Call REPLA(X$, MKS$(CAOPS), 83, 4)
          Call REPLA(X$, MKS$(CAOPS2), 178, 4)
          Call REPLA(X$, Chr$(CATOP%), 87, 1)
          Call REPLA(X$, Chr$(CATOP2%), 177, 1)
          Call REPLA(X$, MKI$(SOLAPAI%), 89, 2)
          Call REPLA(X$, MKS$(LOTESTA), 91, 4)
          Call REPLA(X$, MKS$(HORHERR), 95, 4)
          Call REPLA(X$, MKS$(PIECICL), 99, 4)
          Call REPLA(X$, EQUIP1$, 103, 6)
          Call REPLA(X$, EQUIP2$, 109, 6)
          Call REPLA(X$, EQUIP3$, 115, 6)
          Call REPLA(X$, EQUIP4$, 182, 6)
          Call REPLA(X$, EQUIP5$, 188, 6)
          '
          '\\\OT-06-0046-WAR-10/02/06///
          Call REPLA(X$, HOPFAB1$, 129, 16)
          '\\\OT-06-0046-WAR-FIN///
          '
          Call REPLA(X$, MKI$(PROTERI%), 149, 2)
          Call REPLA(X$, MKD$(PRECTER#), 151, 8)
          Call REPLA(X$, MKI$(MONETER%), 159, 2)
          Call REPLA(X$, MKI$(FEPRETER%), 161, 2)
          
          If NOPY% > ULTOP% Then ULTOP% = NOPY%
          NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
          SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
          If ALTY% > 0 Then
            Mid$(SECO$, 5, 5) = "-" + TRIM$(Str$(ALTY%))
          End If
          List3.AddItem SECO$
          RESECO$(List3.ListCount - 1) = X$  ' Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0)
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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then Exit Sub
   '
   RETEMODI3% = MODIFIC%(3)
   '
   CABAS = LOTESTAN(Ci%)
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
   CONDI$ = " CodIConj = " & Ci% & " AND NumeOper =" & NOPER%
   FEC% = CVINT(VALOBADA("SECOPER", "FECH_REVIS", CONDI$, "NOMESS"))
   '
   If FEC% < 1 Then
        Label60.Caption = ""
        Label62.Caption = ""

        RNC$ = RECFILT$("DATADOPS", 1, MKI$(Ci%))
        URN& = ULTREG&("DATADOPS")
        '
        If Len(RNC$) >= 4 Then
          RENOTA& = CVS(Left$(RNC$, 4))
          If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
              TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
              If CVI(Left$(TTXX$, 2)) = Ci Then
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Then Exit Sub
   '
   ' ACUALIZA REVISION ANTE CUALQUIER CAMBIO EN EL MAESTRO
   If MODIFIC%(0) > 0 Or MODIFIC%(1) > 0 Or MODIFIC%(2) > 0 Or MODIFIC%(3) > 0 Then
        If Control("MASTER", "ACTUREVI") = "SI" Then
            Text25.TEXT = FECHATEX$(HOY(HO$))
        End If
   End If
   
   If MODIFIC%(0) > 0 Then GoSub 90: Call GRAGENITEM
   If MODIFIC%(1) > 0 Then GoSub 90: Call GRAESPECI
   If MODIFIC%(2) > 0 Or (SSTab1.Tab = 1 And GRABATODO% = 1) Then Command9.SetFocus: Call GRAFORMULA
   If MODIFIC%(3) > 0 Then GoSub 90: Call GRAHORUTA
   '
   If GRABATODO% = 1 Then
     If SSTab1.Tab = 3 Then Call GRACOSTAN
   End If
   '

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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Then Exit Sub
   '
   If DERACCE%("GRACOSTA", "Grabar Costos Standard por Articulo") < 2 Then
      Exit Sub
   End If
   '
   RETRA& = Ci
   If RETRA& <= ULTREG&("RECOSCAL") Then
      XX$ = REGLEIDO$("MASTER", RETRA& + 1)
      TIMA% = Asc(Mid$(XX$, 93, 1))
      UMED$ = AJUSTI$(Mid$(XX$, 109, 2), 2)
      FECO% = HOY(HOS$)
      YY$ = String$(128, 0)
      Call REPLA(YY$, MKI$(Ci%), 1, 2)
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Then Exit Sub
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   'SOLICITUD DE DOSIVAC QUE PUEDAN OTORGAR SOLO DERECHO DE MODIFICAR SEGUNDA DESCRIPCION
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If DERACCE%("MODITMAS/NOMESS", "Modificación Datos Maestro de Stock") < 2 Then
      'SI ESTA CONFIGURADA LA DESCRIPCION DOBLE
      If Control("MASTER", "DESLARGA") <> "SI" And InStr(EPAC$, "DOSIVA") > 0 Then
         'VALIDA EL DERECHO
         If DERACCE%("MODI2DES", "Modificación Segunda Descripción") > 1 Then
             Call ACTUREGISTRO("MASTER", "Descridos", "codint = " & Ci%, Left$(Text38, 64), REGAF&, "NOMESS")
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
       If CDIXI% <> Ci Then
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
'            STRSQL = STRSQL & "' where codint= " & CI
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
     If COMALTER%("Se han Modificado Datos del Artículo\'" + TRIM$(CODEXT$(Ci%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(0) = 0
       MODIFIC(0) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   'VALIDACION DE ATRIBUTO A SI REPITE
   If Control("AMASTO", "REPATRIA") = "SI" And TRIM$(Text23.TEXT) <> "" Then
        TTALLE$ = TRIM$(Text23.TEXT)
        CONDI$ = "TALLE = '" & TTALLE$ & "' AND CODINT <> " & Ci
        If CantRegistros("MASTER", CONDI$) > 0 Then
           Call MENSERR(24, "Atributo A :" & TTALLE$ & " Ya Existente en otro Código")
           Text23 = VALOBADA("MASTER", "TALLE", "CODINT = " & Ci%, "NOMESS")
        End If
   End If
   '
   'VALIDACION DE ATRIBUTO A SI REPITE
   If Control("AMASTO", "REPATRIB") = "SI" And TRIM$(Text22.TEXT) <> "" Then
        TCOLOR$ = TRIM$(UCase$(Text22.TEXT))
        CONDI$ = "COLOR = '" & TCOLOR$ & "' AND CODINT <> " & Ci
        If CantRegistros("MASTER", CONDI$) > 0 Then
           Call MENSERR(24, "Atributo B :" & TCOLOR$ & " Ya Existente en otro Código")
           Text22 = VALOBADA("MASTER", "COLOR", "CODINT = " & Ci%, "NOMESS")
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
   strSQL = ("Select * from Master where codint=") & CIXI%
   RMaster.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
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
   TIMA% = XVALO(Text6(0).TEXT)
   RMaster!TIMATE = TIMA%
   '
   FAMI$ = Left$(Text48.TEXT, 2)
   RMaster!familia = FAMI$
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
   XCI = CCOLOR%(TCOLOR$)
   '
   TTALLE$ = TRIM$(Text23.TEXT)
   RMaster!TALLE = Left$(TTALLE, 32)
   XCI = CTALLE%(TTALLE$)
   '
   CO_INTRERNA$ = TRIM$(Text44.TEXT)
   RMaster!Cod_Interna = Left$(CO_INTRERNA$, 64)
   '
   RMaster!Revision = Left$(TRIM$(Text36), 16)
   FREVI% = FECHANUM(Text25.TEXT)
   RMaster!Ferevis = CVDAT(FREVI%)
   '
   STOMINI = XVALO(Text18)
   RMaster!STOMIN = STOMINI
   '
   'GUARDO EN UNA VARIABLE EL VALOR ANTERIOR
   LOTEECOANT = LOTESTAN(Ci%)
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
     
     PREC_ACT# = COSUNI(Ci%)
     COU = XVALO(Text12.TEXT)
     PRE_NUE# = COU
     'SI CAMBIO DE PRECIO GRABA EN NOVEDADES
     If ROUND(COSTO_ANT#, 4) <> ROUND(COU, 4) Then
        Observa$ = "Cambio Desde Maestro de Artículos"
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "J Y M") > 0 Then
            Screen.MousePointer = 1
            OBSERVAC.Caption = "Observación de Cambio de Precio (Mandataria)"
            OBSERVAC.Text28 = ""
            OBSERVAC.Show 1                        'LLAMA A VENTANA DE OBSERVACIONES
            If OBSERVAC.Label1.Caption <> "OK" Or TRIM$(OBSERVAC.Text28) = "" Then
               OBSERVAC.Text28 = ""
               Call COMUNI("Debe Ingresar un Texto Motivo del Cambio de Precio\No se va a Registrar el Mismo")
               GoTo 20
            End If
            Screen.MousePointer = 11
            Observa$ = OBSERVAC.Text28
        End If
        Call ADD_NOVEDAD(CIXI%, XVALO(Text10), PREC_ACT#, PRE_NUE#, Observa$, XVALO(Text13), "COSTO")
     End If
     '
     RMaster!COSUNI = COU
     '
     FECOSTX = Text14.TEXT
     If FECOSTX = "" Then FECOSTX = Now
     On Error Resume Next
     RMaster!FECOSTO = CDate(FECOSTX)
     On Error GoTo 0
     '
   End If
   
20 PORNA = XVALO(Text28.TEXT)
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

'If InStr(EPAC$, "AHP") > 0 Then
'   TRAZ% = 1
'End If
   
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
   RMaster!LETICOM = LTC
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
   
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "codint= " & CIXI%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
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
   
   If Control("MASTER", "FORPATRO") = "SI" And LOTEECOANT <> LOREPO Then
      Call ACTUFORMUPAT(Ci%, LOTEECOANT, LOREPO)
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
   'If CI > 1 Then     ' fuerza actualizacion de descripcion
   '   DDD = DESCRIT$(CI%- 1)
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Or Ci% > NUMAS% Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("MODISPEC", "Modificación de Especificaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(Ci%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       MODIFIC(1) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
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
'   STRSQL = "Update Master set Especificacion = '" & Text43 & "' where codint = " & CI
'   FLEXCONN.Execute STRSQL, cdadreg
   '
   'SE MODIFICA LA GRABACION CON ACTUREGISTRO POR QUE TIENE CAPTURADO EL ERROR QUE ESTA
   'PRODUCIENDO EN LA LINEA DE ARRIBA CUANDO SE INTENTA GRABAR UN TEXTO CON APOSTROFE
   Call ACTUREGISTRO("Master", "Especificacion", "codint = " & Ci%, Text43, REG&)
   '
   Call ACTUREGISTRO("Master", "ArchiPdf", "codint = " & Ci%, Label63, REG&)
   Call ACTUREGISTRO("Master", "Clas_ATC", "codint = " & Ci%, Text81, REG&)
   Call ACTUREGISTRO("Master", "Vers_BC ", "codint = " & Ci%, Text82, REG&)
   Call ACTUREGISTRO("Master", "FeVig_BC", "codint = " & Ci%, FETEXCOR1$(FECHANUM(Text83)), REG&)
   Call ACTUREGISTRO("Master", "FeRev_BC", "codint = " & Ci%, FETEXCOR1$(FECHANUM(Text84)), REG&)
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "codint = " & Ci%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
   End If

   'GRAABAR P.O.E.`S
   '****************
   'MARCA DE BORRADO
   Call ACTUREGISTRO("DOCUREL", "MARBORRA", "COD_INTE = " & Ci%, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To List2.ListCount
      Z$ = List2.List(i& - 1)
      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
10    RNDX$ = RANDSTRING$(8)
      If CantRegistros("DOCUREL", "DESCRIPCION = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10

      Call CreaRegistro("DOCUREL", "DESCRIPCION", RNDX$, "NOMESS")
      Call ACTUREGISTRO("DOCUREL", "COD_INTE", "DESCRIPCION = '" & RNDX$ & "'", Ci%, REG&)
      Call ACTUREGISTRO("DOCUREL", "MARBORRA", "DESCRIPCION = '" & RNDX$ & "'", 0, REG&)
      Call ACTUREGISTRO("DOCUREL", "RUTA", "DESCRIPCION = '" & RNDX$ & "'", RUTAX$, REG&)
      Call ACTUREGISTRO("DOCUREL", "DESCRIPCION", "DESCRIPCION = '" & RNDX$ & "'", DESCRI$, REG&)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCUREL", "MARBORRA = 1 AND COD_INTE = " & Ci%, REGAF&, "NOMESS")
   '
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "COD_INTE = " & Ci%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "DOCUREL", "MINEPARTS", "DOCUREL", CONDI$)
   End If
   
   Screen.MousePointer = 1
    
   
   MODIFIC(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAFORMULA()
   '
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Or Ci% > NUMAS% Then Exit Sub
   If YACAR%(2) < 1 Then Exit Sub
   If MODIFIC%(2) < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Fórmula del Artículo\'" + TRIM$(CODEXT$(Ci%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(2) = 0
       MODIFIC(2) = 0
       Call CARGENITEM
       GoTo 99
     End If
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
   End If
   
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
   For i& = 1 To CAIT%
     EXPLOLIN(i&) = List1.List(i& - 1)
     CICOMPO% = CODINT%(Left$(EXPLOLIN(i&), 16))
     Call EXPLOMUL(CICOMPO%)
     For JJI% = 1 To CONTAMUL%
       If CIK%(JJI%) = Ci Then
         Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(Ci%)) + "\" + DESCRIT0$(Ci%) + ".")
         YACAR%(2) = 0
         MODIFIC(2) = 0
         SSTab1.Tab = 1
         Call CARGENITEM
         GoTo 99
       End If
     Next JJI%
   Next i&
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Or Ci% > NUMAS% Then Exit Sub
   If YACAR%(3) < 1 Then MODIFIC%(3) = 0
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If DERACCE%("MODIRUTA", "Modificación Hojas de Ruta") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(Ci%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       MODIFIC%(3) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If

   
   Screen.MousePointer = 11
   
   Call ABRESECOP 'llama a abreconexion
   'sql
   Dim RSECOPER As ADODB.Recordset
   Set RSECOPER = New ADODB.Recordset
   'RSECOPER.CursorType = adOpenDynamic
   RSECOPER.CursorType = adOpenKeyset
   RSECOPER.LockType = adLockPessimistic
   strSQL = ("Select * from Secoper where CodIConj=") & Ci
   RSECOPER.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
   With RSECOPER
      strsql_U = ("Update Secoper set StatOper = -1 where StatOper>0 and CodIConj=") & Ci
      FLEXCONN.Execute strsql_U, cdadreg
   
   
   '
   'With Secoper
   '  YAPRIMO% = 0
   '  On Error Resume Next
   '  .MoveFirst
   '  If Err > 0 Then GoTo 65
   '  On Error GoTo 0
     '
   '  ABUS$ = "CodIConj =" & CI%& " AND StatOper > 0"
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
     '.FindFirst "CODICONJ =" & CI%& " and NumeOper = " & NOPYI%
     'If .NoMatch = False Then
     Dim MIRS As ADODB.Recordset
     strsql_S = ("Select * from Secoper where CODICONJ=") & Ci% & " and NumeOper=" & NOPYI%
     Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql_S))
     If Not (MIRS.EOF And MIRS.BOF) Then
       CANPRE = MIRS!NUOPER_PREP
       CATPRE% = MIRS!CATOPER_PREP
     End If
     'End If
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CODIEMPR = CodiEmp%
     On Error GoTo 0
     !CODICONJ = Ci
     !CODXCONJ = CODEXT$(Ci%)
     !NUMEOPER = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
     !AltOper = Asc(Mid$(RESECO$(LL% - 1), 88, 1))
     !descoper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
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
     !CATEGOPS = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
     !CATEGOPS2 = Asc(Mid$(RESECO$(LL% - 1), 177, 1))
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
     !Fech_Revis = CVDAT(FE%)
     !USUA_REVIS = USERNAME$
     
     .Update
   
   Next LL%
   RSECOPER.Close
   Set RSECOPER = Nothing
   End With
      strsql_D = ("Delete Secoper where StatOper=-1 and CodIConj=") & Ci
      FLEXCONN.Execute strsql_D, cdadreg
   
   
   '
63 '.FindFirst "CODICONJ =" & CI
   'If .NoMatch = False Then
   '  If !statoper = (-1) Then
   '    .Delete
   '    GoTo 63
   '  End If
   'End If
   '
64 '.FindNext "CODICONJ =" & CI
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
Sub GRADATOPER(Ci%, NOPE%, NHPF$)
   '
   If Ci% > 0 Then
      If Ci <= NUMAS% Then
         URN& = ULTREG&("DATADOPS")
         RNC$ = RECFILT$("DATADOPS", 1, MKI$(Ci%))
         KK& = (-3)
         '
         TTYY$ = REGBLAN$("DATADOPS")
         Call REPLA(TTYY$, MKI$(Ci%), 1, 2)
         Call REPLA(TTYY$, MKI$(NOPE%), 3, 2)
         Call REPLA(TTYY$, NHPF$, 5, 16)
         'Call REPLA(TTYY$, Text2.TEXT, 21, 4)
         FEXI% = HOY(HOS$)
         Call REPLA(TTYY$, MKI$(FEXI%), 25, 2)
         Call REPLA(TTYY$, MKI$(FEXI%), 27, 2)
         'Call REPLA(TTYY$, Text5.TEXT, 29, 8)
         USUA$ = USERNAME$ 'TRIM$(Text6(0).TEXT): If USUA$ = "" Then USUA$ = USERNAME$
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
               Call GRAREG("DATADOPS", MKI$(Ci%) + MKI$(0) + Space$(124), RENOTA&)
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
      If SSTab1.Tab = 4 Then
        If TRIM$(Label28) = "" Then Call LIMPIARSTOCKS
        If Label15(2) = Label28 Then
           Exit Sub
        Else
           Label15(2) = Label28
'           Call Label15_Change(2)
        End If
      End If
      '
      If SSTab1.Tab = 5 Then
        'HISTORICO DE MOVIMIENTOS
        If TRIM$(Label28) = "" Then MSF(3).Rows = 1
        If Label15(8) = Label28 Then
           Exit Sub
        Else
            Label15(8) = Label28
'           Call Label15_Change(2)
'           Call CARGAR_STOCKS
'           Call CARGAR_STOCKS_PROVEEDOR(CI1%)
        End If
      End If
      If SSTab1.Tab = 6 Then
        'ESTADISTICAS
        If TRIM$(Label28) = "" Then MSF(5).Rows = 1: MSF(7).Rows = 1: MSF(6).Rows = 1: MSF(8).Rows = 1
        If Label15(10) = Label28 Then
           Exit Sub
        Else
            Label15(10) = Label28
        End If
      End If

      Call CARDATITEM
      '
   End If
End Sub
Sub CARGA_ESTADISTICAS(CI1%)
    Call CARGA_ENCABEZADO(Me.MSF(5), "Presup./f7.0;Cliente/A22;Proveedor/A12;Costo U$S/F10.2;M./A3;P.Unit./F10.2;Fecha/D8;NCLIEN/F0", Me)
    Call CARGA_ENCABEZADO(Me.MSF(7), "Factura./A20;Cliente/A24;M./A3;P.Unit./F10.2;Fecha/D8", Me)
    Call CARGA_ENCABEZADO(Me.MSF(6), "Referencia/A12;U.Cotiz./F8", Me)
    Call CARGA_ENCABEZADO(Me.MSF(8), "Referencia/A12;U.Vendidas/F8", Me)
    '
    MSF(5).Rows = 1: MSF(7).Rows = 1: MSF(6).Rows = 1: MSF(8).Rows = 1
    
       If CI1 > 0 Then
          Call HISTORIALPRESUPUESTOS(MSF(5), CI1%, 0)
          Call HISTORIALVENTAS(MSF(7), CI1%, 0)
          Call HISTORIAL_COTIZA(MSF(6), CI1%)
          Call HISTORIAL_VENTAS(MSF(8), CI1%)
       End If
End Sub
Sub HISTORIAL_VENTAS(MSF As msFlexGrid, CI1%)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    Call TRACE(20, 10, 40)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT90%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_90 = SALE - INGRE: If ULT_90 < 1 Then ULT_90 = 0
    '
    Call TRACE(20, 20, 40)
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT180%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_180 = SALE - INGRE: If ULT_180 < 1 Then ULT_180 = 0
    '
    Call TRACE(20, 30, 40)
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT365%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_365 = SALE - INGRE: If ULT_365 < 1 Then ULT_365 = 0
    '
    MSF.Rows = 4: MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    Call TRACE(20, 1, 1)
    '
    
End Sub

Sub HISTORIALVENTAS(MSF As msFlexGrid, CI1%, NC)

     MSF.Rows = 1 ': MSF.Cols = 5
     CONDI$ = "COD_INTE =  " & CI1
     CONDI$ = CONDI$ + "  AND CODIMOVI = 'RT'"
     If NC > 0 Then CONDI$ = CONDI$ + " AND  Nume_Clie =  " & NC
     FIN& = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
     
     SQLX$ = "SELECT  DOSECLAR, VALOUNIT,Nume_Clie,MONEVAL, FECHMOV, TIPOCAM "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE " & CONDI$
     SQLX$ = SQLX$ + " group by DOSECLAR, VALOUNIT,Nume_Clie,MONEVAL, FECHMOV, TIPOCAM "
     SQLX$ = SQLX$ + " ORDER BY DOSECLAR DESC "
     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = SAFETEXT$(!DOSECLAR)
        MSF.TextMatrix(J&, 2) = rasocli$(XVALO(!NUME_CLIE))
        MSF.TextMatrix(J&, 3) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
        '
'        MON% = XVALO(!MONEVAL)
'        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
'        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
'           TIPOCAMFACTURA# = XVALO(VALOBADA("CAJABANC", "COTIZA_DOLAR", "CODICOM_LAR = '" & MSF.TextMatrix(j&, 1) & "'", "NOMESS"))
''           TIPOCAMBIODEPRESUP# = XVALO(!TIPOCAM) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
''           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
''               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMBIODEPRESUP# ' CALCULA
''           Else
'               'PRECIOUNITARIO# = XVALO(!VALOUNIT) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
'               'PARA QUE SIEMPRE MUESTRE EL VALOR EN DOLARES
'               If Abs(TIPOCAMFACTURA#) < 0.005 Then TIPOCAMFACTURA# = TICAMONE(2)
'               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMFACTURA#
'               On Error GoTo 0
'               MON% = 2
''           End If
'        End If
        ''
        'se busca el dolar en la factura y en el remito por el caso que solo se realiza remito y no factura
        MON% = XVALO(!MoneVal)
        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!TIPOCAM) 'TOMA EL CAMBIO EN QUE FUE GUARDADO EN MOVISTO
           If TIPOCAMBIODEPRESUP# > 0.0005 And TIPOCAMBIODEPRESUP# <> 1 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMBIODEPRESUP# ' CALCULA
               MON% = 2
           Else
               'SI NO TIENE EL TIPO DE CAMBIO EN EL MOVISTO LO BUSCA EN LA FACTURA
               TIPOCAMFACTURA# = XVALO(VALOBADA("CAJABANC", "COTIZA_DOLAR", "CODICOM_LAR = '" & MSF.TextMatrix(J&, 1) & "'", "NOMESS"))
               If TIPOCAMFACTURA# > 0.0005 And TIPOCAMFACTURA# <> 1 Then
                 PRECIOUNITARIO# = XVALO(!VALOUNIT) / TIPOCAMFACTURA#
               Else
                 PRECIOUNITARIO# = XVALO(!VALOUNIT) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               End If
               MON% = 2
           End If
        End If

        '
        MSF.TextMatrix(J&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        MSF.TextMatrix(J&, 5) = TRIM$(FECHATEX$(CVINT(!FechMov)))
        '
        .MoveNext
       Loop
     End With
     Call numerargrilla(MSF, 0, 1)
End Sub
Sub HISTORIAL_COTIZA(MSF As msFlexGrid, CI1%)
    Call TRACE(20, 10, 40)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT90%) & "' AND CodiInt = " & CI1
    ULT_90 = XVALO(VALOBADA("PRESUVEN", "SUM(Cantit)", CONDI$, "NOMESS"))
    '
    
    Call TRACE(20, 20, 40)
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT180%) & "' AND CodiInt = " & CI1
    ULT_180 = XVALO(VALOBADA("PRESUVEN", "SUM(Cantit)", CONDI$, "NOMESS"))
    '
    Call TRACE(20, 30, 40)
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT365%) & "' AND CodiInt = " & CI1
    ULT_365 = XVALO(VALOBADA("PRESUVEN", "SUM(Cantit)", CONDI$, "NOMESS"))
    '
    MSF.Rows = 4
    MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    Call TRACE(20, 1, 1)
    '
End Sub
Sub GENKARDEXAHP(Optional MOSTRARCOSTOS%)
   '
   Call ABRECONEXION
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
      Call CARDEPOS
      PRIDEP% = COSU%(1)
   End If
   '
'   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   'PARA LIMPIAR LA GRILLA POR MAS QUE EL CODIGO NO SEA VALIDO

   iRow& = 1: MSF(3).Rows = 15
   MSF(3).Rows = 1
    MSF(4).Rows = 1
'   For I& = 1 To MSF(3).Rows - 1
'     MSF(3).Row = I
'      For j& = 0 To MSF(3).Cols - 1
'          MSF(3).COL = j
'         MSF(3).TextMatrix(I&, j&) = ""
'         MSF(3).CellBackColor = RGB(255, 255, 255)
'         MSF(3).CellForeColor = RGB(0, 0, 0)
'      Next j&
'   Next I&
   
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = 0
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text6(1)): Desde = FETEXCOR1$(DESFE%)
   ' DESDE = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text6(2)): Hasta = FETEXCOR1$(HASFE%)
   ' HASTA = CDbl(CVDAT(HASFE%))
   If DESFE% > HASFE% Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   If DESFE% = 0 Then Exit Sub
   INCLUVIE% = 0:   'If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD = CIXX%
   condi0$ = "Cod_Inte = " & CIDD & " "
   If DEPXX% > 0 Then
      If DEPXX% = PRIDEP% Then
          condi0$ = condi0$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
        Else
          condi0$ = condi0$ + "AND DepoMovi = " & DEPXX% & " "
      End If
   End If
   '
   Screen.MousePointer = 11

   Call ABRESTOCKS
   iRow& = 0: MSF(3).Rows = 1
'   For I& = 1 To GRID.Rows - 1
'      For j& = 0 To GRID.Cols - 1
'         GRID.TextMatrix(I&, j&) = ""
'      Next j&
'   Next I&
   ACUIN = 0: ACUSA = 0
   If INCLUVIE% = 1 Then
      condi0$ = condi0$ + "AND CodiMovi <> 'KU' "
      ACUIN = XVALO(VALOBADA("MOVIVIE", "SUM(CANTINGRE)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
      ACUSA = XVALO(VALOBADA("MOVIVIE", "SUM(CANTSALID)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   End If
   ACUIN = ACUIN + XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   ACUSA = ACUSA + XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", "FECHMOV < '" & Desde & "' AND " & condi0$, "NOMESS"))
   '

   SARRAS# = ACUIN - ACUSA
   If Abs(SARRAS#) >= 0.05 Then
       iRow& = iRow& + 1
       If iRow& > MSF(3).Rows - 1 Then
          MSF(3).Rows = iRow& + 1
       End If

       FEX = DIANTE(DESFE%)
       MSF(3).TextMatrix(iRow&, 1) = FECHATEX$(FEX)
       REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
       MSF(3).TextMatrix(iRow, 6) = REFEM$
       CASALDO# = SARRAS#
       MSF(3).TextMatrix(iRow, 10) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
       iRow& = iRow& + 1  ' PARA QUE DEJE UNO LIBRE
   End If

   JJ& = 0
   ' SARRAS# = 0
   INCLUPOS% = 0
'   If CIDD% > 0 Then
'     Call ABRESTOCKS
'     Call ABRESTOCKVIE
'     If INCLUVIE% = 1 Then
'     '
'        SQLX1$ = " SELECT * FROM MOVIVIE WITH(NOLOCK) "
'        SQLX1$ = SQLX1$ + "WHERE Cod_Inte = " & CIDD% & " "
'        If DEPXX% > 0 Then
'          If DEPXX% = PRIDEP% Then
'              SQLX1$ = SQLX1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
'            Else
'              SQLX1$ = SQLX1$ + "AND DepoMovi = " & DEPXX% & " "
'          End If
'        End If
'        SQLX1$ = SQLX1$ + "AND CodiMovi <> 'KU' "
        SQLX1$ = SQLX1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
   BASEAC$ = "MOVISTO": If INCLUVIE% = 1 Then BASEAC$ = "MOVIVIE"
   '
   MSF(3).Redraw = False
   FIN& = CantRegistros&(BASEAC$, condi0$, "NOMESS")
20 SQLX1$ = "SELECT FECHMOV, CodiMovi, DOPRILAR,  DoSecLar, IdenLote, ReferCor,  DEPOMOVI,  CANTINGRE, CANTSALID,ValoUnit,NUME_CLIE"
   SQLX1$ = SQLX1$ + " FROM " + BASEAC$ + " WITH(NOLOCK)"
   SQLX1$ = SQLX1$ + " WHERE " + condi0$
   SQLX1$ = SQLX1$ + " AND FECHMOV >= '" & Desde & "' "
   SQLX1$ = SQLX1$ + " AND FECHMOV <= '" & Hasta & "' "
   SQLX1$ = SQLX1$ + " ORDER BY FechMov , FeReMovi ASC "
'   Dim MOVIVIETMP As ADODB.Recordset
'   Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
   Set MOVIVIETMP = READSET(SQLX1$)
   With MOVIVIETMP
      Do While Not .EOF
              If CVINT(!FechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
         IDLOT$ = SAFETEXT$(!idenlote)
         If XVALO(Label8) > 0 Then
            If PROVELOTE%(IDLOT$) <> XVALO(Label8) Then GoTo 65
         End If
             '
         iRow& = iRow& + 1
         Call TRACE(20, iRow&, FIN&)
         
         If iRow& > MSF(3).Rows - 1 Then
            MSF(3).Rows = iRow& + 1
         End If
         MSF(3).TextMatrix(iRow, 1) = FECHATEX$(CVINT(!FechMov))
         MSF(3).TextMatrix(iRow&, 2) = SAFETEXT$(!CODIMOVI)
         MSF(3).TextMatrix(iRow&, 3) = TRIM$(SAFETEXT$(!DOPRILAR))
         MSF(3).TextMatrix(iRow&, 4) = SAFETEXT$(!DOSECLAR)
            IDLOT$ = SAFETEXT$(!idenlote)
         MSF(3).TextMatrix(iRow&, 5) = IDLOT$
         MSF(3).TextMatrix(iRow&, 6) = SAFETEXT$(!ReferCor)
         MSF(3).TextMatrix(iRow&, 7) = XVALO(!depomovi)
            CAINGRE# = XVALO(!CANTINGRE)
            CASALI# = XVALO(!CANTSALID)
            CASALDO# = CASALDO# + XVALO(!CANTINGRE) - XVALO(!CANTSALID)
         MSF(3).TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         MSF(3).TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         MSF(3).TextMatrix(iRow&, 10) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
         If MOSTRARCOSTOS% > 0 Then
         'ESTA RUTINA POR EL LOTE ME DEVUELVE LOS VALORES DE LOS PARAMETROS SEGUN EL LOTE
         FOB# = 0: FobNacionalizado# = 0: PrexVenta# = 0: NPROV = 0: VIAENT% = 0: MONCOMP% = 0
         Call FobPorcgastoPreventaProveedVia(IDLOT$, CIXX%, FOB#, FobNacionalizado#, PrexVenta#, NPROV, VIAENT%, MONCOMP%)
         VIA_ENTREGAX$ = TRIM$(ECOARCH$("VIASENTR", MKI$(VIAENT%)))
         'SI ES FACTURA TRAE EL PRECIO DE VENTA DEL  MOVISTO
         If Left(MSF(3).TextMatrix(iRow&, 4), 3) = "FC-" Then
           PrexVenta# = XVALO(!VALOUNIT)
           CONDI$ = "Codicom_lar = '" & TRIM$(MSF(3).TextMatrix(iRow&, 4)) & "'"
           PrexVenta# = PrexVenta# / COTIZACIONCABANC(TRIM$(MSF(3).TextMatrix(iRow&, 4)), XVALO(!NUME_CLIE))
         End If
         If FOB# > 0 Then MSF(3).TextMatrix(iRow&, 11) = TRIM$(FORMATNUM$(FOB#, "F8.2"))
         If FobNacionalizado# > 0 Then MSF(3).TextMatrix(iRow&, 12) = TRIM$(FORMATNUM$(FobNacionalizado#, "F8.2"))
         If PrexVenta# > 0 Then MSF(3).TextMatrix(iRow&, 13) = TRIM$(FORMATNUM$(PrexVenta#, "F8.2"))
         If NPROV > 0 Then MSF(3).TextMatrix(iRow&, 14) = TEL2CLI$(-1 * NPROV)
         If VIAENT% > 0 Then MSF(3).TextMatrix(iRow&, 15) = VIA_ENTREGAX$
         End If
'            Call CompletaCombo(IDLOT$)
65    .MoveNext
      Loop
   End With
   MOVIVIETMP.Close
   Set MOVIVIETMP = Nothing
   If BASEAC$ = "MOVIVIE" Then
      BASEAC$ = "MOVISTO"
      GoTo 20
   End If
   '
   MSF(3).Redraw = True
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = MSF(3).Rows - 12: If TPXI& < 1 Then TPXI& = 1
   On Error Resume Next
   MSF(3).TopRow = TPXI&
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   Call TRACE(20, 1, 1)
   '
End Sub




Sub CARGAR_STOCKS_PROVEEDOR(CI1%, Optional MOSTRARCOSTOS%)
    CODII% = CI1%
    If CODII% < 1 Then
      Call COMUNI("Código Inexistente o No Válido")
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Campos$ = "Cta.Prov/F5;R.Social/A12;Stock/F6.1;V.Entrega/A8;Ubicación/A20;Ult.F.Recep/D7;Idenlote/A0;Cto Fob/F8.2;Nacionalizado/F12.2;%/F0;P.Venta-Form./F0;P.Venta Propio/F0"
    Call CARGA_ENCABEZADO(MSF(1), Campos$, Me)
    '
    'EN ESTA CONSULTA TRAIGO TODO LO QUE SE INGRESO POR LOTE Y SALIO POR LOTE
    MSF(1).Rows = 1
    J& = 0
    ACUMLA_SALDO = 0
    'ORIGINAL
'    SQLX$ = " SELECT B.COD_EXTE, B.NUMPROV, M.COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO M with(nolock)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP B ON M.IDENLOTE = B.IDENLOTE"
'    SQLX$ = SQLX$ + " WHERE B.COD_INTE = " & CODII%
'    SQLX$ = SQLX$ + " AND M.COD_INTE = " & CODII%
'    SQLX$ = SQLX$ + " GROUP BY M.Cod_Inte, B.NumProv,B.COD_EXTE"
'    'SQLX$ = SQLX$ + " ORDER BY M.FECHMOV"

    SQLX$ = " SELECT B.COD_EXTE, B.NUMPROV, M.COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO, B.VIA_ENTREGA,M.IDENLOTE FROM MOVISTO M with(nolock)"
    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP B ON M.IDENLOTE = B.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE B.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " AND M.COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY M.Cod_Inte, B.NumProv,B.COD_EXTE, B.VIA_ENTREGA,M.IDENLOTE "
    
    Set rs1 = READSET(SQLX$)
    With rs1
        Do While Not .EOF
             NPROV = XVALO(!NumProv)
             NprovReal = NPROV
             NP& = NPROV
             If ProveedorRealExterior%(NP&) > 0 Then
               LOTE$ = SAFETEXT$(!idenlote)
               InfoProveMineparts% = CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, Ret_ProvExterior, Ret_ViaEntExt, Ret_PjeNacion, Ret_CoefNacion)
               NprovReal = Ret_ProvExterior
             End If

             RAZONPROVEEDOR$ = rasocli$(-1 * NprovReal)
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             If MOSTRARCOSTOS% > 0 Then
                 LOTE$ = ULTIMOLOTERECIBIDO$(CODII%, NPROV)
                 VIALOTEXX% = XVALO(!Via_Entrega)
                 VIA_ENTREGAX$ = TRIM$(ECOARCH$("VIASENTR", MKI$(XVALO(!Via_Entrega))))
                 
                 'COSTO FOB
                 COSFOB# = COSTOFOB#(LOTE$)
                 MONECOMP% = MONELOTE%(LOTE$)
                 FEINGLOTE% = FECHAINGLOTE%(LOTE$)
                 COEFCON = 1
                 If MONECOMP% <> 2 Then
                    'ESTO ES PARA QEU CUANDO SE COMPRA EN $ TENGA EL TIPO DE CAM DE LA OC Y LO PASE A U$S
                    TipoCambioCom = TipoCambioCompra(LOTE$)
                    If TipoCambioCom > 0 Then
                       COEFCON = TICAMONE(MONECOMP%) / TipoCambioCom
                    Else
    '                   COEFCON = TICAMONE(MONECOMP%) / TICAMONE(MONEMI%)
                       COEFCON = TICAMONE(MONECOMP%) / 9.42 'ultimo valor antes de antes de guardar el valor de la cotizadcion en la o.c
                    End If
                 End If
                 COSFOB# = COEFCON * COSFOB#
                 KOSTO$ = FORMATNUM$(COSFOB#, "F12.2")
                 '
                 
                 NUMLOTE& = XVALO(Mid$(LOTE$, 4))
                 PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NUMLOTE&, , MONECOMP%)
                 GASNAC# = XVALO(COSFOB# * PORCEN_GASNAC#)
                 MODONACIONALIZACION% = 0
    
                 If GASNAC# <= 0 Then
                   GASNAC# = COSFOB#
                   MODONACIONALIZACION% = 1
                 End If
                 

                 PREVENTAX# = PRECIOVENTA_X_FORMULA(XVALO(GASNAC#), VIALOTEXX%, CODII%, MODONACIONALIZACION%, NPROV)
             End If
             J& = J& + 1
             MSF(1).Rows = J& + 1
             MSF(1).TextMatrix(J&, 1) = NprovReal
             MSF(1).TextMatrix(J&, 2) = RAZONPROVEEDOR$
             MSF(1).TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             MSF(1).TextMatrix(J&, 4) = VIA_ENTREGAX$
             If NPROV > 0 Then
                MSF(1).TextMatrix(J&, 5) = UBISTOSQL$(CODII%, NPROV)
             Else
                MSF(1).TextMatrix(J&, 5) = UBISTO$(CODII%)
             End If
             MSF(1).TextMatrix(J&, 6) = FECHATEX$(FEINGLOTE%)
             If MOSTRARCOSTOS% > 0 Then
                MSF(1).TextMatrix(J&, 7) = TRIM$(LOTE$)
                MSF(1).TextMatrix(J&, 8) = TRIM$(FORMATNUM$(KOSTO$, "F10.2"))
                MSF(1).TextMatrix(J&, 9) = TRIM$(FORMATNUM$(GASNAC#, "F10.2"))
                MSF(1).TextMatrix(J&, 10) = TRIM$(FORMATNUM$(PORCENT_ENTRE_2_VALORES(PREVENTAX#, (GASNAC#)), "F9.2"))
                MSF(1).TextMatrix(J&, 11) = TRIM$(FORMATNUM$(PREVENTAX#, "F10.2"))
                  VENTAAHPEDITADO = PRECIODEVENTA_AHP(CODII%, NPROV, VIALOTEXX%)
                MSF(1).TextMatrix(J&, 12) = TRIM$(FORMATNUM$(VENTAAHPEDITADO, "F10.2"))
             End If
' CAMPOS$ = "Cta.Prov/F8;R.Social/A20;Stock/F9.1;Ubicación/A36;Ult.F.Recep/D10;Idenlote/A0(6);Cto Fob/F10.2;Cto.Nacion./F10.2;%/F5;P.Venta-Form./F10.2;P.Venta Propio/"
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    '
    'AGREGO AL FINAL EL ACUMULADO DE LOS MOVIMIENTOS DEL ITEM QUE SE GENERARON SIN LOTE POR EJEMPLO FACTURAS GENERADAS SIN LOTE.
    SQLX$ = " SELECT COD_INTE, SUM(CANTSALID) AS SALEN, SUM(CANTINGRE) AS INGRESO FROM MOVISTO With(nolock)"
    SQLX$ = SQLX$ + " WHERE (IDENLOTE IS NULL OR IDENLOTE = '') "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CODII%
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte"
    Set rs1 = READSET(SQLX$)
    With rs1
        Do While Not .EOF
             SALDOX = XVALO(!ingreso) - XVALO(!SALEN)
             '
             J& = J& + 1
             MSF(1).Rows = J& + 1
             MSF(1).TextMatrix(J&, 1) = "0"
             MSF(1).TextMatrix(J&, 2) = "No Identificado"
             MSF(1).TextMatrix(J&, 3) = FORMATNUM$((XVALO(SALDOX)), "F12.1")
             '
             ACUMLA_SALDO = ACUMLA_SALDO + SALDOX
            .MoveNext
        Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    Screen.MousePointer = 1
    
'    FRMZELECHO.Caption = "Código: " & CODXX$ & "     -     Saldo: " & ACUMLA_SALDO
'    FRMZELECHO.Show 1
'
'    PROVED = XVALO(RESP_ZELE_VECT(1))
'    If PROVED > 0 Then Call VERUBICA(CODII%, PROVED)
    

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
'   FORSELEC.ABUSQUE = Text1
   Call Command9_Click
'   Call SELECHO("MASTER")
'   COD$ = TRIM$(VALACT1$("MASTER"))
'   If COD$ <> "" Then
'     Text1.TEXT = COD$
'   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
End Sub

Private Sub Text10_Change()
   Label33 = rasocli$((-1) * XVALO(Text10))
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
     Label41.Caption = Unimed$(CICOMPO%)
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
     Label41.Caption = Unimed$(CICOMPO%)
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
         Label41.Caption = Unimed$(CIE%)
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

Private Sub Text6_Change(Index As Integer)
   If Index = 0 Then
        VALONU = XVALO(Text6(0).TEXT)
        If VALONU < 0 Or VALONU > 255 Then
           Text6(0).TEXT = "": Exit Sub
        End If
        Label29.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
        MODIFIC%(0) = 1
    ElseIf Index = 1 Or Index = 2 Then
        Call Label15_Change(2)
    End If
End Sub

Private Sub Text6_DblClick(Index As Integer)
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6(0).TEXT = VALONUE$
   End If
End Sub

Private Sub Text16_Change()
'   VALONU = XVALO(Text16.TEXT)
'   If VALONU < 0 Or VALONU > ULTREG("GRUCOVEN") Then
'      Text16.TEXT = "": Exit Sub
'   End If
'   Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(Int(VALONU)))
'   MODIFIC%(0) = 1

   'SE MODIFICA POR QUE NO SIEMPRE COINCIDE LA CANTIDAD DE REGISTROS CON EL CODIGO DE FILA.
   
   VALONU = XVALO(Text16.TEXT)
   
   If VALONU < 0 Or ECOARCH$("GRUCOVEN", MKI$(Int(VALONU))) = "" Then
      Label31.Caption = "": Exit Sub
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
           For i& = 1 To FIN&
              Call TRACE(20, i&, FIN&)
              X$ = REGLEIDO$("MASTER", i& + 1)
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
                 Call GRAREG("MASTER", X$, i& + 1)
              End If
           Next i&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For i& = 1 To FIN&
              Call TRACE(20, i&, FIN&)
              X$ = REGLEIDO$("INVERT", i& + 1)
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
                 Call GRAREG("INVERT", X$, i& + 1)
              End If
           Next i&
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
                CIX% = II&
                Call REPLA(X$, MKI$(CIX%), 1, 2)
                Call GRAREG("MASTER", X$, II& + 1)
            End If
            COD$ = AJUSTI$(TRIM$(Mid$(X$, 77, 15)), 15) + Mid$(X$, 92, 1)
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
        X$ = REGBLAN$("INVERT")
        Call REPLA(X$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", X$, 1)
        '
        J& = 1
        FIN& = NV% - 1
        For II& = 0 To FORMALTER.LINOSORT.ListCount - 1
            Call TRACE(20, II&, FIN&)
            FORMALTER.LINOSORT.ListIndex = II&
            CODISEL$(II&) = Mid$(FORMALTER.LINOSORT.TEXT, 2, 16)
            CISEL%(II&) = XVALO(Mid$(FORMALTER.LINOSORT.TEXT, 18))
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
                m$ = REGLEIDO$("MASTER", RGM&)
                Call REPLA(m$, CODNUE$, 77, 16)
                Call GRAREG("MASTER", m$, RGM&)
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
        k% = 0
        FIN& = NV%
        For XL& = 1 To NV%
             Call TRACE(20, XL&, FIN&)
             IK$ = REGLEIDO$("INVERT", XL& + 1)
             If Mid$(IK$, 16, 1) <> Chr$(96) Then
               If TRIM$(Left$(IK$, 16)) <> "" Then
                 '
                 Ci% = CVI(Mid$(IK$, 17, 2))
                 DE$ = DESCRIT$(Ci%)
                 '
                 EC$ = Space$(62) + MKI$(Ci%)
                 Call REPLA(EC$, Left$(IK$, 16), 1, 16)
                 Call REPLA(EC$, DE$, 17, 46)
                 '
                 k% = k% + 1: ECM$ = EC$: Put #FIECO%, k%, ECM$
                              CCT$ = MKI$(Ci%): Put #FICIA%, k% + 1, CCT$
                 '
               End If
             End If
        Next XL&
        '
        CCT$ = MKI$(k%): Put #FICIA%, 1, CCT$
        '
        For K1% = k% + 1 To LOF(FIECO%) / 64
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
        ENCU% = 1: Ci% = CVI(Mid$(IK$, 17, 2)): RGI& = L&
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
       DESCRI$ = TRIM$("Lt.Std." + TRIM$(Str$(LOSTAND)) + " " + TRIM$(Unimed$(CI0%)))
       DESCRI$ = " (" + DESCRI$ + ")"
       LOMADE% = 60 - Len(DESCRI$)
       If LOMADE% > 0 Then
          DESCRI$ = Left$(TRIM$(DESCRIT0$(CI0%)), LOMADE%) + DESCRI$
       End If
       '
       Y$ = REGBLAN$("SECOPSEL")
       Call REPLA(Y$, MKI$(CI0%), 1, 2)
       Call REPLA(Y$, DESCRI$, 3, 60)
       v& = v& + 1
       Call GRAREG("SECOPSEL", Y$, v&)
       '
       Call REPLA(Y$, Space$(60), 3, 60)
       v& = v& + 1
       Call GRAREG("SECOPSEL", Y$, v&)
       '
       Do Until .EOF = True
         NOPY% = !NUMEOPER
         ALTY% = !AltOper
         DOPE$ = !descoper
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
         CATOP% = !CATEGOPS
         CATOP2% = !CATEGOPS2
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
         Call REPLA(Y$, NOPES$, 3, 8)
         Call REPLA(Y$, DOPES$, 11, 46)
         Call REPLA(Y$, EQUIPO$, 57, 6)
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
         TSTANBA = MINVARI    ' TIESTAN(!COD_INTE, !NumeOper, 1, 1, ALTOP%)
         CADEN = 0
         If TSTANBA > 0 Then CADEN = 60 / TSTANBA
         Call REPLA(Y$, MKS$(CADEN), 125, 4)
         '
         v& = v& + 1
         Call GRAREG("SECOPSEL", Y$, v&)
         .MoveNext
       Loop
       '
     End With
     Set SECOPTEMP = Nothing
     '
     Call REPLA(Y$, String$(126, 0), 3, 126)
     v& = v& + 1
     Call GRAREG("SECOPSEL", Y$, v&)
     '
     Y$ = REGBLAN$("SECOPSEL")
     v& = v& + 1
     Call GRAREG("SECOPSEL", Y$, v&)
     '
25 MasterTemp.MoveNext
   U& = U& + 1
   Loop
   '
   Call SETULTREG("SECOPSEL", v&)
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
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      CIBAS% = i&
      '
      Call AREXPLO(CIBAS%, 1, 3)
      If CAIT% > 0 Then CIANU%(CIBAS%) = 1
      For J% = 1 To CAIT%
        CIANU%(CIJ%(J%)) = 1 '\\\OT-06-0078-WAR-03/03/06/// - BANDERA QUE E COMPONENTE DE ALGUNA ESTRUCTURA
        For k% = J% + 1 To CAIT%
          If CIJ%(J%) = CIJ%(k%) Then
            COD1$ = TRIM$(CODEXT$(CIBAS%))
            COD2$ = TRIM$(CODEXT$(CIJ%(J%)))
            If CODINT%(COD1$) = CIBAS% Then ' SUPRIME ITEMS DE BAJA
              CAERR% = CAERR% + 1
              MENSA$ = "Repetición del Código " + COD2$ + "\en Fórmula de Codigo " + COD1$
              If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
            End If
          End If
        Next k%
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
89  Next i&
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
      Y$ = REGBLAN$("COMPOBAJ")
      JJ& = 0
      For J% = 1 To NMS%
        If CIANU%(J%) = 2 Then
          Call REPLA(Y$, MKI$(J%), 1, 2)
          JJ& = JJ& + 1
          Call GRAREG("COMPOBAJ", Y$, JJ&)
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
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      Ci% = i&
      If STATUIT(Ci%) > 0 Then
        ACTIX% = 0
        CODYY$ = CODEXT$(Ci%)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            PORMILLAR% = 0
            Call LEEEXPLO(CODYY$, 1)
            If CAIT% > 0 Then
              For k% = 1 To CAIT%
                CI1% = CIJ%(k%)
                CAN = USOI(k%)
                PD% = PDI%(k%)
                If CI1% > 0 Then
                  If CI1% <= NUMAS% Then
                    UM$ = Unimed$(CI1%)
                    Call CONVERUNID(CAN, UM$)
                    If Abs(CAN) > 0.00005 Then
                      Z$ = Z00$
                      Y$ = Y00$
                      '\\\OT-05-0702-WAR-01/11/05///
                      v$ = V00$
                      '\\\OT-05-0702-WAR-FIN///
                      Call REPLA(Z$, MKI$(Ci%), 1, 2)
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
                      Call REPLA(v$, MKI$(Ci%), 1, 2)
                      Call REPLA(v$, MKI$(CI1%), 3, 2)
                         ALI1$ = NUPLANO$(CI1%)
                      Call REPLA(v$, ALI1$, 5, 24)
                         ALI1$ = ""
                      Call REPLA(v$, UM$, 29, 2)
                         ATR1$ = ATRIB1$(CI1%)
                      Call REPLA(v$, ATR1$, 31, 32)
                         ATR2$ = ATRIB2$(CI1%)
                      Call REPLA(v$, ATR2$, 63, 32)
                      Call REPLA(v$, MKS$(CAN), 95, 4)
                      Call REPLA(v$, Chr$(PD%), 99, 1)
                         REFE$ = Mid$(EXPLOLIN$(k%), 129)
                      Call REPLA(v$, REFE$, 109, 64)
                      '\\\OT-06-0352-MH-30/08/06///
                      If STATITEM%(CI1%) = -1 Then
                         Call REPLA(v$, "Baja", 173, 4)
                      End If
                      '\\\OT-06-0352-MH-FIN///
                      '\\\OT-05-0702-WAR-FIN///
                      '
                      Call REPLA(Y$, MKI$(Ci%), 1, 2)
                      Call REPLA(Y$, MKI$(CI1%), 3, 2)
                      Call REPLA(Y$, UM$, 5, 2)
                        PESOMET = PESMETRO(CI1%)
                      Call REPLA(Y$, MKS$(PESOMET), 7, 4)
                        'LARGO CORTE (MM)
                        LACORT = XVALO(Mid$(EXPLOLIN$(k%), 111, 6))
                      Call REPLA(Y$, MKS$(LACORT), 11, 4)
                      Call REPLA(Y$, MKS$(CAN), 19, 4)
                        'CANTIDAD
                        CANTIUN = XVALO(Mid$(EXPLOLIN$(k%), 89, 9))
                      Call REPLA(Y$, MKS$(CANTIUN), 15, 4)
                      Call REPLA(Y$, Chr$(PD%), 11, 1)
                        'REFERENCIA
                        REFE$ = Mid$(EXPLOLIN$(k%), 129)
                      Call REPLA(Y$, REFE$, 33, 64)
                      '\\\OT-06-0352-MH-30/08/06///
                      If STATITEM%(CI1%) = -1 Then
                         Call REPLA(Y$, "Baja", 97, 4)
                      End If
                      '\\\OT-06-0352-MH-FIN///
                      
                      '
                      J& = J& + 1
                      Call GRAREG("LIESTRU", Z$, J&)
                      Call GRAREG("LISTRULA", Y$, J&)
                      '\\\OT-05-0702-WAR-01/11/05///
                      Call GRAREG("LISTRUAT", v$, J&)
                      '\\\OT-05-0702-WAR-FIN///
                      ACTIX% = 1
                      '\\\OT-05-0663-WAR-FIN///
                    End If
                  End If
                End If
              Next k%
            End If
          End If
        End If
        '
        If ACTIX% > 0 Then
          Z$ = Z00$
          Call REPLA(Z$, MKI$(Ci%), 1, 2)
          J& = J& + 1
          Call GRAREG("LIESTRU", Z$, J&)
          '\\\OT-05-0663-WAR-07-10-05///
          Y$ = Y00$
          Call REPLA(Y$, MKI$(Ci%), 1, 2)
          Call GRAREG("LISTRULA", Y$, J&)
          '\\\OT-05-0663-WAR-FIN///
          '\\\OT-05-0702-WAR-01/11/05///
          v$ = V00$
          Call REPLA(v$, MKI$(Ci%), 1, 2)
          Call GRAREG("LISTRUAT", v$, J&)
          '\\\OT-05-0702-WAR-FIN///
        End If
        '
      End If
      '
    Next i&
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
   Ci% = CODINT%(Text1.TEXT)
   If Ci% < 1 Then Exit Sub
   '
60 Screen.MousePointer = 11
   IMPLOSI07.List1.Clear
   IMPLOSI07.Caption = "Implosión " + DESCRIT$(Ci%)
   IMPLOSI07.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI07.Left = Left + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left + (Width - ScaleWidth) / 2
   'IMPLOSI07.Show
   '
   'Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas WHERE FORMULAS!CODIELEM=" & CI%, dbOpenDynaset, dbReadOnly)
   Dim FORMULASTEMP As ADODB.Recordset
   Set FORMULASTEMP = New ADODB.Recordset
   FORMULASTEMP.CursorType = adOpenKeyset
   FORMULASTEMP.LockType = adLockReadOnly
   SQLX$ = "SELECT codiconj, codielem, usobruto, codint, descripcion, codigo, status, umedida FROM FORMULAS INNER JOIN MASTER ON codiconj=codint WHERE codielem=" & CStr(Ci%) & " AND status > 0"
   'SQLX$ = "SELECT * FROM Formulas WHERE CODIELEM=" & CI
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
      USOX = FORMULASTEMP!usobruto
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
          Call REPLA(Z$, SAFETEXT$(FORMULASTEMP!Codigo), 1, 16)
          Call REPLA(Z$, SAFETEXT$(FORMULASTEMP!Descripcion), 17, 32)
          Call REPLA(Z$, SAFETEXT$(FORMULASTEMP!UMEDIDA), 51, 2)
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
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    REGQ& = 0:
    REGR& = 0
    '
    OPX% = COMALTER%("Formato de Salida:\\Normal\Con Referencias")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 2 Then
       OPXST% = COMALTER%("Formato de Salida:\\Solo Referencias\Referencias + Stock")
    End If
    Screen.MousePointer = 11
    Dim TIESTRU%(32767), ESPARSTRU%(32767)
    '
    NMS% = NUMAS%
    FIN& = NUMAS%
    J& = 0
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      Ci% = i&
      If STATUIT(Ci%) > 0 Then
        CODYY$ = CODEXT$(Ci%)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            PORMILLAR% = 0
            Call LEEEXPLO(CODYY$, 1)
            If CAIT% > 0 Then
              TIESTRU%(Ci%) = 1
              For k% = 1 To CAIT%
                CI1% = CIJ%(k%)
                If CI1% > 0 Then
                  If CI1% <= NMS% Then
                    ESPARSTRU%(CI1%) = 1
                  End If
                End If
              Next k%
            End If
          End If
        End If
      End If
    Next i&
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
      Ci% = MasterTemp!CODINT
      If (TIESTRU%(Ci%) = 1 And ESPARSTRU%(Ci%) = 0) Or CI1% = CI2% Then
        Call EXPLOMUL(Ci%)
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
            UME$ = Unimed$(CIK%(II%))
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
                   If OPXST% = 2 Then
                     CI1% = CIK%(II%)
                     AA = STOCKACT(CI1%)
                     STOCK1$ = FORMATNUM$(STOCKACT(CI1%), "F12." & CNTDC$(CI1%))
                     Call REPLA(YY$, STOCK1$, 157, 12)
                   End If
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
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
       Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
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
    Ci% = MasterTemp!CODINT
    If STATUIT(Ci%) = 1 Or CI1% = CI2% Then
      PORMILLAR% = 0
      Call EXPLOCOM(Ci%, 1)
      If CONTAMUL% > 0 Then
        '
        ACTIX% = 0
        For k% = 1 To CONTAMUL%
          CIX1% = CIK%(k%)
          CAN = CANT(k%)
          PD% = 0
          If CIX1% > 0 Then
            If CIX1% <= NUMAS% Then
              UM$ = Unimed$(CIX1%)
              Call CONVERUNID(CAN, UM$)
              If Abs(CAN) > 0.0005 Then
                Z$ = REGBLAN$("LIESTRU")
                Call REPLA(Z$, MKI$(Ci%), 1, 2)
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
        Next k%
        '
        If ACTIX% > 0 Then
          Z$ = REGBLAN$("LIESTRU")
          Call REPLA(Z$, MKI$(Ci%), 1, 2)
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
     n1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #n1%, AA$
     '
10   If EOF(n1%) Then
       Close #n1%
       GoTo 15
     End If
     '
     Line Input #n1%, AA$
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
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #n1%
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
     n1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #n1%, AA$
     '
60   If EOF(n1%) Then
       Close #n1%
       GoTo 65
     End If
     '
     Line Input #n1%, AA$
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
         Uso = XVALO(ENTE$) + (XVALO(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR1$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If Uso > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
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
     For i& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", i&)
       Ci% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(Ci%)
   Next i&
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
     n1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #n1%, AA$
     '
10   If EOF(n1%) Then
       Close #n1%
       GoTo 19
     End If
     '
     Line Input #n1%, AA$
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
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #n1%
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
     n1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #n1%, AA$
     '
60   If EOF(n1%) Then
       Close #n1%
       GoTo 65
     End If
     '
     Line Input #n1%, AA$
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
         Uso = XVALO(ENTE$) + (XVALO(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR1$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If Uso > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For i& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", i&)
       Ci% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(Ci%)
   Next i&
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
   Ci% = CODINT%(CODD$)
   If Ci% < 1 Then Exit Sub
   CICON% = Ci
   '
   Screen.MousePointer = 11
   Call EXPLOMUL(Ci%)
   If CONTAMUL% < 2 Then Exit Sub
   '
   IMPLOSI07.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI07.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI07.List1.Clear
   IMPLOSI07.Caption = "Explosión Multinivel " + DESCRIT$(Ci%)
   IMPLOSI07.Show
   '
   COEFCA = 1: If PORMILLAR% = 1 Then COEFCA = 1000
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
     Call REPLA(Z$, SP$ + DESCRIT0$(CIK%(KKI%)), 17, 32)
     Call REPLA(Z$, Unimed$(CIK%(KKI%)), 51, 2)
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
      X$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(X$, 2))
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
    X$ = REGLEIDO$("TABLATMU", J&)
    If COTMU$ = Mid$(X$, 1, 4) Then
      MSTMU# = CVD(Mid$(X$, 37, 8)) 'Duración en mseg.
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
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   CI1% = CVI(Left$(X$, 2))
   COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2))
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
      CIXI% = !cod_inte
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
         Call MUESTRADOC1(ARFULLNAME$)
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
Sub MUESTRADOC1(ARFULLNAME$)
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

Private Sub TXTBUSCAR_Change()
    '
    Dim MSF_A As msFlexGrid
    Dim MSF_B As msFlexGrid
    If TRIM$(TXTBUSCAR) = "" Then
      Set MSF_A = MSF(0)
      Set MSF_B = MSF(2)
      
      MSF_B.Visible = False
      MSF_A.Visible = True
      Exit Sub
    End If
    '
'    If TRIM$(TXTBUSCAR) = "" Then
'      Set MSF_A = MSF(3)
'      Set MSF_B = MSF(4)
'
'      MSF_A.Visible = False
'      MSF_B.Visible = True
'      Exit Sub
'    End If
    
    MSF(2).Rows = 2
    For i& = 1 To MSF(2).Cols - 1
       MSF(2).TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF(2).Visible = True
    MSF(2).ZOrder 0
    MSF(2).Cols = MSF(0).Cols
    MSF(2).Redraw = False
    MSF(0).Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To MSF(0).Rows - 1
      For J& = 1 To MSF(0).Cols - 1
        If (InStr(UCase$(MSF(0).TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF(0).TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF(2).Rows = REG& + 1
            For H& = 1 To MSF(2).Cols - 1
              MSF(2).TextMatrix(REG&, H&) = MSF(0).TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF(2).Redraw = True
    
    '
End Sub

Sub HISTORIALPRESUPUESTOS(MSF As msFlexGrid, CI1%, NC)
     MSF.Rows = 1
     CONDI$ = "CODIINT = " & CI1%
     If XVALO(NC) > 0 Then CONDI$ = CONDI$ + " AND NroClie =  " & NC
     FIN& = CantRegistros&("PRESUVEN", CONDI$, "NOMESS")
     
     SQLX$ = "SELECT  NroPres, RasoClie, PreUnit,NroClie,MonedEmi,COTIZAMONEDA, FechEmis, Num_Prov, COSTO_UNIT"
     SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE " & CONDI$
'     If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND NroClie =  " & NC
     SQLX$ = SQLX$ + " group by NroPres, RasoClie, PreUnit,NroClie,MonedEmi,COTIZAMONEDA, FechEmis, Num_Prov, COSTO_UNIT"
     SQLX$ = SQLX$ + " ORDER BY NroPres desc"
     Set rs = READSET(SQLX$)
     With rs
       MSF.Redraw = False
       Do While Not .EOF
         
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        MSF.Rows = J& + 1
        MSF.TextMatrix(J&, 1) = XVALO(!NroPres)
'        MSF.TextMatrix(j&, 2) = RASOCLI$(XVALO(!NROCLIE))
        MON% = XVALO(!MonedEmi)
        PRECIOUNITARIO# = XVALO(!PreUnit) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        COSTOUNITARIO# = XVALO(!COSTO_UNIT)
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZAMONEDA) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 And TIPOCAMBIODEPRESUP# <> 1 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TIPOCAMBIODEPRESUP# ' CALCULA
               COSTOUNITARIO# = XVALO(!COSTO_UNIT) / TIPOCAMBIODEPRESUP# ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
           End If
           MON% = 2
        End If
        MSF.TextMatrix(J&, 3) = TEL2CLI$(-1 * XVALO(!Num_Prov))
        MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(COSTOUNITARIO#, "F12.2"))
        MSF.TextMatrix(J&, 5) = SIMBOLPES$(MON%)
        MSF.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
          FE% = CVINT(!FECHEMIS)
        MSF.TextMatrix(J&, 7) = FECHATEX$(FE%)
        MSF.TextMatrix(J&, 8) = XVALO(!NROCLIE)
        .MoveNext
       Loop
     End With
     '
     For i& = 1 To MSF.Rows - 1
        NC = XVALO(MSF.TextMatrix(i&, 8))
        MSF.TextMatrix(i&, 2) = rasocli$(NC)
     Next i&
     Call numerargrilla(MSF, 0, 1)
     MSF.Redraw = True
End Sub

