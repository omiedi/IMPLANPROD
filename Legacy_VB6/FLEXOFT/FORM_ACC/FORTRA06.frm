VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FORTRA06 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Trabajo - A-B-M-L Base de Datos de O.T."
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FORTRA06.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   225
      Left            =   105
      TabIndex        =   134
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   330
      Left            =   105
      TabIndex        =   133
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text1 
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
      Left            =   1155
      TabIndex        =   84
      Text            =   " "
      Top             =   210
      Width           =   1725
   End
   Begin VB.TextBox Text3 
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
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   83
      Top             =   525
      Width           =   2925
   End
   Begin VB.TextBox Text8 
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
      Left            =   6090
      TabIndex        =   82
      Top             =   840
      Width           =   2340
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
      Left            =   9660
      TabIndex        =   81
      Top             =   210
      Width           =   1125
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
      Left            =   10755
      Picture         =   "FORTRA06.frx":0442
      TabIndex        =   80
      Top             =   525
      Width           =   175
   End
   Begin VB.TextBox Text11 
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
      Left            =   9135
      TabIndex        =   79
      Top             =   840
      Width           =   1650
   End
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
      Height          =   285
      Left            =   9660
      TabIndex        =   78
      Top             =   525
      Width           =   1125
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
      Left            =   5145
      TabIndex        =   77
      Top             =   525
      Width           =   3300
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
      Left            =   2835
      Picture         =   "FORTRA06.frx":074C
      TabIndex        =   76
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text18 
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
      Left            =   5145
      TabIndex        =   75
      Top             =   210
      Width           =   3300
   End
   Begin VB.TextBox Text7 
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
      Left            =   5145
      TabIndex        =   74
      Top             =   840
      Width           =   870
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8D8C0&
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
      Height          =   880
      Left            =   8400
      TabIndex        =   30
      Top             =   7800
      Visible         =   0   'False
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "FORTRA06.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   250
         Width           =   550
      End
      Begin VB.CommandButton Command1 
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
         Left            =   155
         Picture         =   "FORTRA06.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
      End
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
         Left            =   725
         Picture         =   "FORTRA06.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command8 
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
         Picture         =   "FORTRA06.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Acceso Directo a otras Aplicaciones"
         Top             =   3685
         Width           =   650
      End
      Begin VB.CommandButton Command15 
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
         Height          =   690
         Left            =   105
         Picture         =   "FORTRA06.frx":26DA
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Guardar "
         Top             =   1630
         Width           =   640
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
         Picture         =   "FORTRA06.frx":29E4
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Copia de Especificación de Venta"
         Top             =   3000
         Width           =   650
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "FORTRA06.frx":2B2E
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ver  y Re-Imprimir"
         Top             =   4380
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   7245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
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
         Picture         =   "FORTRA06.frx":2E38
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Imprime Ficha de Especificación de Venta"
         Top             =   5835
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Picture         =   "FORTRA06.frx":34A2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Emisión de Listados"
         Top             =   5145
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Picture         =   "FORTRA06.frx":37AC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6525
         Width           =   650
      End
      Begin VB.CommandButton Command99 
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
         Picture         =   "FORTRA06.frx":3BEE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "FORTRA06.frx":3EF8
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exit"
         Top             =   105
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
         Height          =   690
         Left            =   105
         Picture         =   "FORTRA06.frx":4042
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
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
         Picture         =   "FORTRA06.frx":434C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alta de Nueva Especificación de Venta"
         Top             =   2310
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FORTRA06.frx":4496
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   150
      TabIndex        =   5
      Top             =   1260
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   327681
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Especificacion de Venta"
      TabPicture(0)   =   "FORTRA06.frx":63B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Horas / Observaciones"
      TabPicture(1)   =   "FORTRA06.frx":63D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "Frame10"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Lista de Grupos"
      TabPicture(2)   =   "FORTRA06.frx":63F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame10 
         Caption         =   "Resumen de Horas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   -74790
         TabIndex        =   112
         Top             =   4935
         Width           =   3585
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1335
            TabIndex        =   121
            Top             =   1155
            Width           =   945
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2415
            TabIndex        =   120
            Top             =   1155
            Width           =   945
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Horas:"
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
            Left            =   -630
            TabIndex        =   119
            Top             =   1215
            Width           =   1905
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1335
            TabIndex        =   118
            Top             =   735
            Width           =   945
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2415
            TabIndex        =   117
            Top             =   735
            Width           =   945
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Previstas:"
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
            Left            =   -630
            TabIndex        =   116
            Top             =   795
            Width           =   1905
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1335
            TabIndex        =   115
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2415
            TabIndex        =   114
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Indirectas:"
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
            Left            =   -630
            TabIndex        =   113
            Top             =   480
            Width           =   1905
         End
      End
      Begin VB.Frame Frame9 
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
         Height          =   6000
         Left            =   -71010
         TabIndex        =   71
         Top             =   525
         Width           =   6510
         Begin VB.TextBox Text35 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5400
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   420
            Width           =   6270
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Horas Previstas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   -74790
         TabIndex        =   62
         Top             =   2520
         Width           =   3585
         Begin VB.TextBox Text32 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   69
            Top             =   1320
            Width           =   945
         End
         Begin VB.TextBox Text30 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   65
            Top             =   420
            Width           =   945
         End
         Begin VB.TextBox Text29 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   64
            Top             =   720
            Width           =   945
         End
         Begin VB.TextBox Text28 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   63
            Top             =   1020
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   6
            Left            =   2415
            TabIndex        =   111
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   5
            Left            =   2415
            TabIndex        =   110
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   109
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   108
            Top             =   1020
            Width           =   945
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Left            =   -630
            TabIndex        =   107
            Top             =   1845
            Width           =   1905
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2415
            TabIndex        =   106
            Top             =   1785
            Width           =   945
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1335
            TabIndex        =   105
            Top             =   1785
            Width           =   945
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mont.Externo:"
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
            Left            =   -630
            TabIndex        =   70
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Caldereria:"
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
            Left            =   -645
            TabIndex        =   68
            Top             =   465
            Width           =   1905
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mecanizado:"
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
            Left            =   -630
            TabIndex        =   67
            Top             =   750
            Width           =   1905
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mont.Interno:"
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
            Left            =   -630
            TabIndex        =   66
            Top             =   1065
            Width           =   1905
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Horas Indirectas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   -74790
         TabIndex        =   47
         Top             =   525
         Width           =   3585
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   60
            Top             =   1020
            Width           =   945
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   58
            Top             =   720
            Width           =   945
         End
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
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
            Left            =   1335
            TabIndex        =   48
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Left            =   -630
            TabIndex        =   104
            Top             =   1530
            Width           =   1905
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2415
            TabIndex        =   103
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1335
            TabIndex        =   102
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   101
            Top             =   1020
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   100
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   99
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Cal:"
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
            Left            =   -630
            TabIndex        =   61
            Top             =   1065
            Width           =   1905
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Avor:"
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
            Left            =   -630
            TabIndex        =   59
            Top             =   750
            Width           =   1905
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ingenieria:"
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
            Left            =   -645
            TabIndex        =   49
            Top             =   465
            Width           =   1905
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   6090
         Left            =   -74895
         TabIndex        =   16
         Top             =   420
         Width           =   10410
         Begin VB.TextBox Text15 
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
            Left            =   4425
            MaxLength       =   2
            TabIndex        =   131
            Top             =   165
            Width           =   840
         End
         Begin VB.TextBox Text6 
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
            Left            =   7050
            TabIndex        =   129
            Top             =   165
            Width           =   1125
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
            Left            =   8145
            Picture         =   "FORTRA06.frx":640C
            TabIndex        =   128
            Top             =   165
            Width           =   175
         End
         Begin VB.CommandButton Command7 
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
            Height          =   330
            Left            =   9240
            TabIndex        =   127
            Top             =   130
            Width           =   1065
         End
         Begin VB.Frame Frame15 
            Caption         =   "Frame15"
            Height          =   5580
            Left            =   7560
            TabIndex        =   126
            Top             =   595
            Width           =   15
         End
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            Height          =   5580
            Left            =   2415
            TabIndex        =   125
            Top             =   595
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   5580
            Left            =   1680
            TabIndex        =   124
            Top             =   595
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   5580
            Left            =   945
            TabIndex        =   123
            Top             =   595
            Width           =   15
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   5450
            Left            =   0
            ScaleHeight     =   5385
            ScaleWidth      =   10350
            TabIndex        =   17
            Top             =   600
            Width           =   10410
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5100
               ItemData        =   "FORTRA06.frx":6716
               Left            =   -30
               List            =   "FORTRA06.frx":6718
               TabIndex        =   122
               Top             =   315
               Width           =   10410
            End
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   10515
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   " Grupo       Prog.     Comp.   Denominación / Descripción                                             Observaciones               "
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
                  Left            =   0
                  TabIndex        =   19
                  Top             =   50
                  Width           =   10530
               End
            End
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisión:"
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
            Left            =   2100
            TabIndex        =   132
            Top             =   210
            Width           =   2280
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Revisión:"
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
            Left            =   4725
            TabIndex        =   130
            Top             =   210
            Width           =   2280
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6120
         Left            =   200
         TabIndex        =   8
         Top             =   420
         Width           =   10320
         Begin VB.Frame Frame3 
            Caption         =   "Condiciones de Pago"
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
            Left            =   0
            TabIndex        =   41
            Top             =   4515
            Width           =   4800
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   900
               IntegralHeight  =   0   'False
               ItemData        =   "FORTRA06.frx":671A
               Left            =   120
               List            =   "FORTRA06.frx":671C
               TabIndex        =   56
               Top             =   600
               Width           =   4530
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00FFFFC0&
               Height          =   345
               Left            =   120
               ScaleHeight     =   285
               ScaleWidth      =   4470
               TabIndex        =   54
               Top             =   315
               Width           =   4530
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "     It.     Descripcion                     Importe         Fecha"
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
                  Left            =   0
                  TabIndex        =   55
                  Top             =   20
                  Width           =   4950
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "L.Entrega"
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
            Left            =   4935
            TabIndex        =   39
            Top             =   4515
            Width           =   2955
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
               Height          =   1170
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   315
               Width           =   2745
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Otros"
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
            Left            =   7980
            TabIndex        =   11
            Top             =   4515
            Width           =   2325
            Begin VB.TextBox Text21 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               TabIndex        =   4
               Text            =   " "
               Top             =   1155
               Width           =   870
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
               Index           =   5
               Left            =   1995
               TabIndex        =   3
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   5
               Left            =   1410
               TabIndex        =   2
               Text            =   " "
               Top             =   315
               Width           =   630
            End
            Begin VB.Label DETEXT 
               BackColor       =   &H00E0E0E0&
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
               Index           =   5
               Left            =   105
               TabIndex        =   14
               Top             =   630
               Width           =   2040
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Comision % :"
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
               Left            =   -525
               TabIndex        =   13
               Top             =   1230
               Width           =   1710
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Representante:"
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
               Left            =   -345
               TabIndex        =   12
               Top             =   375
               Width           =   1710
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10305
            Begin VB.CommandButton Command6 
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
               Left            =   9765
               Picture         =   "FORTRA06.frx":671E
               TabIndex        =   51
               Top             =   420
               Width           =   175
            End
            Begin VB.CommandButton Command5 
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
               Left            =   4830
               Picture         =   "FORTRA06.frx":6A28
               TabIndex        =   50
               Top             =   420
               Width           =   175
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
               Height          =   285
               Left            =   5985
               TabIndex        =   36
               Top             =   420
               Width           =   2040
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
               Left            =   8715
               TabIndex        =   35
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text24 
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
               Left            =   3780
               TabIndex        =   1
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text25 
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
               Left            =   1050
               TabIndex        =   0
               Top             =   420
               Width           =   2040
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Oferta:"
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
               Left            =   4920
               TabIndex        =   38
               Top             =   470
               Width           =   1065
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   8085
               TabIndex        =   37
               Top             =   470
               Width           =   585
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   3150
               TabIndex        =   34
               Top             =   470
               Width           =   585
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido:"
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
               Left            =   -890
               TabIndex        =   10
               Top             =   470
               Width           =   1905
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -945
            TabIndex        =   42
            Top             =   840
            Width           =   11670
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00E0E0E0&
               Height          =   2955
               Left            =   9555
               ScaleHeight     =   2895
               ScaleWidth      =   1560
               TabIndex        =   96
               Top             =   525
               Width           =   1620
               Begin VB.TextBox Text38 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   98
                  Top             =   -30
                  Width           =   1650
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   375
                  Left            =   0
                  TabIndex        =   97
                  Top             =   420
                  Width           =   1575
               End
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               Height          =   2955
               Left            =   1020
               ScaleHeight     =   2895
               ScaleWidth      =   795
               TabIndex        =   94
               Top             =   525
               Width           =   855
               Begin VB.TextBox Text36 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   95
                  Top             =   -30
                  Width           =   850
               End
            End
            Begin VB.TextBox Text37 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
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
               Left            =   0
               TabIndex        =   73
               Text            =   "Text36"
               Top             =   0
               Width           =   825
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFC0&
               Height          =   330
               Left            =   1020
               ScaleHeight     =   270
               ScaleWidth      =   10095
               TabIndex        =   52
               Top             =   210
               Width           =   10155
               Begin VB.Line Line24 
                  X1              =   0
                  X2              =   10080
                  Y1              =   315
                  Y2              =   315
               End
               Begin VB.Line Line23 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line2 
                  X1              =   840
                  X2              =   840
                  Y1              =   315
                  Y2              =   0
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   $"FORTRA06.frx":6D32
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
                  Left            =   50
                  TabIndex        =   53
                  Top             =   30
                  Width           =   10200
               End
               Begin VB.Line Line3 
                  X1              =   860
                  X2              =   860
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line22 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
               End
            End
            Begin VB.TextBox Text17 
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
               Left            =   1890
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   46
               Top             =   525
               Width           =   7680
            End
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   7
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   6
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de E.V.:"
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
      Left            =   0
      TabIndex        =   93
      Top             =   315
      Width           =   1140
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   -210
      TabIndex        =   92
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   4020
      TabIndex        =   91
      Top             =   630
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   4200
      TabIndex        =   90
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   8925
      TabIndex        =   89
      Top             =   315
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrega:"
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
      Left            =   8805
      TabIndex        =   88
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   8610
      TabIndex        =   87
      Top             =   945
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   3990
      TabIndex        =   86
      Top             =   315
      Width           =   1110
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   330
      Left            =   3360
      TabIndex        =   85
      Top             =   105
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueOT 
         Caption         =   "Nuevo O.T."
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Consultas y Listados"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiCli 
      Caption         =   "Edición"
      Begin VB.Menu CoInEv 
         Caption         =   "Copia Interactiva de E.V."
      End
      Begin VB.Menu CoInGr 
         Caption         =   "Copia Interactiva de Grupos"
      End
      Begin VB.Menu php5a 
         Caption         =   "-"
      End
      Begin VB.Menu BusCli 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FORTRA06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

Private Sub AyuCli_Click()
  Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
End Sub

Private Sub BusCli_Click()
Call Command20_Click
End Sub

Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub CoInEv_Click()
    If SSTab1.Tab = 0 Or SSTab1 = 1 Then
        Call Command22_Click
    Else
        Call COMUNI("Para realizar esta Copia debe Situarse en la Zolapa de E.V. o en la de Horas/Observaciones")
    End If
End Sub

Private Sub CoInGr_Click()
    If SSTab1.Tab = 2 Then
        Call Command22_Click
    Else
        Call COMUNI("Para realizar esta Copia debe Situarse en la Zolapa de Grupos.")
    End If
End Sub

Private Sub Command10_Click()
 If TRIM(Text1) <> "" Then
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text6.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text6.TEXT)
 End If
End Sub

Private Sub Command11_Click()
   '
 '  Call COMUNI("No Disponible en la presente versión")
'   If Val(TRIM$(Text1)) > 0 Then
'     Call SELECHO("DEUDOR1")
'     NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
'     If NCLIE% < 1 Then Exit Sub
'     '
'     NC% = NCLIE%
'     Text2 = TRIM$(Str$(NC%))
'     Text4 = RASOCLI$(NC%)
'   End If
   '
End Sub

Private Sub Command12_Click()
   '
   Call ABREORTRA
   '
   Screen.MousePointer = 11
   With ListaMat
10   .MoveFirst
     '
20   NUOTR$ = !Nro_Otrab
     NUGRU$ = !Nro_Grup
     NUOR% = !NuOrden
     CANTX = !CANTID
     DECOR$ = !Descri_Cor
     NUPLA$ = !NPlano
     NUPOS$ = !Posic
     MATER$ = !Material
     CODIEX$ = !CODEXTE
     CLAX$ = !CLAVE
     '
     SQLX$ = "SELECT * FROM LISTAMAT "
     SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
     SQLX$ = SQLX$ + "AND Nro_Grup = '" & NUGRU$ & "' "
     SQLX$ = SQLX$ + "AND NuOrden = " & NUOR% & " "
     SQLX$ = SQLX$ + "AND CANTID = " & CANTX & " "
     SQLX$ = SQLX$ + "AND Descri_Cor = '" & DECOR$ & "' "
     SQLX$ = SQLX$ + "AND NPlano = '" & NUPLA$ & "' "
     SQLX$ = SQLX$ + "AND Posic = '" & NUPOS$ & "' "
     SQLX$ = SQLX$ + "AND Material = '" & MATER$ & "' "
     SQLX$ = SQLX$ + "AND Codexte = '" & CODIEX$ & "' "
     SQLX$ = SQLX$ + "AND Clave = '" & CLAX$ & "' "
     Set LIMAGRU = OTrabajo.OpenRecordset(SQLX$, 4)
     '
     LIMAGRU.MoveLast
     CAREGIS& = LIMAGRU.RecordCount
     If CAREGIS& < 2 Then
       If .EOF Then
         .Close
         Call CompactDatabase(LUDAT$ + "\BDACCESS\OTRABAJO.MDB", LUDAT$ + "\BDACCESS\OTRABAJNUE.MDB")
         Call COMUNI("La Base de Datos Ha Sido recuperada\como 'OTRABAJNUE.MDB'.")
         Screen.MousePointer = 1
         Stop
         End
       End If
       .MoveNext
       GoTo 20
     End If
     '
     ABUS$ = "Nro_Otrab = '" & NUOTR$ & "' "
     ABUS$ = ABUS$ + "AND Nro_Grup = '" & NUGRU$ & "' "
     ABUS$ = ABUS$ + "AND NuOrden = " & NUOR% & " "
     ABUS$ = ABUS$ + "AND CANTID = " & CANTX & " "
     ABUS$ = ABUS$ + "AND Descri_Cor = '" & DECOR$ & "' "
     ABUS$ = ABUS$ + "AND NPlano = '" & NUPLA$ & "' "
     ABUS$ = ABUS$ + "AND Posic = '" & NUPOS$ & "' "
     ABUS$ = ABUS$ + "AND Material = '" & MATER$ & "' "
     ABUS$ = ABUS$ + "AND Codexte = '" & CODIEX$ & "' "
     ABUS$ = ABUS$ + "AND Clave = '" & CLAX$ & "' "
     '
     CONTA = 0
30  .FindNext ABUS$
     If .NoMatch = False Then
       .Delete
       CONTA = CONTA + 1
       GoTo 30
     End If
     '
     GoTo 10
     '
   End With
   '
End Sub

Private Sub Command16_Click()
If TRIM(Text1) <> "" Then
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text10.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text10.TEXT)
End If
End Sub


Private Sub Command22_Click()
   '
   If Text1 = " " Then
    Call COMUNI("Debe Seleccionar una Especificación de Venta.")
    Exit Sub
   End If
   '
   If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
     If DERACCE%("ALTNUEXCO", "Alta de Nueva Orden de Trabajo por Copia") < 2 Then
          Exit Sub
       End If
       '
9      Screen.MousePointer = 1
       '
       'FALTA PONER EL CODIGO PARA QUE TRAIGA LAS FECHAS POR DEFAULT
       '
       VDEF$ = Space$(64) + String$(6, 0)
       '
       OBj$ = OBJPLA$("DEACODEV", VDEF$)
       If OBj$ = "" Then GoTo 99
       VDEF$ = OBj$
       '
       NUOTR1$ = TRIM$(Mid$(OBj$, 1, 16))
       NUOTR2$ = TRIM$(Mid$(OBj$, 17, 48))
       '//////////////VALIDACIONES////////////
       If TRIM$(NUOTR1$) = "" Then
          Call COMUNI("Número de Orden de Trabajo no válida")
          GoTo 9
       End If
       '
       If TRIM$(NUOTR2$) = "" Then
          Call COMUNI("Número de Orden de Trabajo no válida")
          GoTo 9
       End If
       '
       Call ABREORTRA
       OrdeTrab.FindFirst "Nro_Otrab = '" & NUOTR1$ & "' "
    '   If OrdeTrab.NoMatch = False Then
    '      Call COMUNI("Numero de Orden de Trabajo no Válido o Pre-Existente")
    '      GoTo 9
    '   End If
    '   '
       With OrdeTrab
        USUA$ = !uSUCONF_OTRAB
        USUC$ = !UsuCont_Otrab
        USUAP$ = !UsuApro_Otrab
        NOC$ = !NroOcom_Otrab
        FPCLIP$ = !FeOcom_Otrab
        NUOFER$ = !NroOfer_Otrab
        FEOFER$ = !FeOfer_Otrab
        CAOPE# = !CanPed_Otrab
        ESPECI$ = !Especi_oTRAB
        MONEOTR% = !mone_Otrab
        COMOTR# = !CotiMone_Otrab
        LUENOTR$ = !LENTREGA_OTRAB
        VEOTR% = !VENDED_OTRAB
        COVEOTR$ = !ComiVend_Otrab
        HSINGE% = !HsInge_Otrab
        HSAVOR% = !HsAvor_Otrab
        HSCCAL% = !HsCCal_Otrab
        HSCAL% = !HsCald_Otrab
        HSMEC% = !HsMeca_Otrab
        HSMOIN% = !HsMoInt_Otrab
        HSMOEXT% = !HsMoExt_Otrab
        HSINGINF% = !HsIngeInf_Otrab
        HSAVIN% = !HsAvorInf_Otrab
        HSCCINF% = !HsCCalInf_Otrab
        HSCAINF% = !HsCaldInf_Otrab
        HSMEINF% = !HsMecaInf_Otrab
        HSMOINF% = !HsMointInf_Otrab
        HSMOEINF% = !HsMoExtInf_Otrab
        OBSOTR$ = !Observa_Otrab
        PREVEOTR# = !preven_Otrab
       End With
       '
       Call ABREORTRA
       OrdeTrab.FindFirst "Nro_Otrab = '" & NUOTR2$ & "' "
        If OrdeTrab.NoMatch = False Then
       With OrdeTrab
          .Edit
    '        MsgBox !NRO_OTRAB
            !uSUCONF_OTRAB = USUA$ 'usuario de confección
            !UsuCont_Otrab = USUAC$ 'usuario que controlo
            !UsuApro_Otrab = USUAP$ 'usuario que aprobara
            !NroOcom_Otrab = NOC$ 'numero de orden de compra del cliente
            !FeOcom_Otrab = FPCLIP$  'fecha de pedido del cliente
            !NroOfer_Otrab = NUOFER$ 'Numero de Oferta
            !FeOfer_Otrab = FEOFER$ 'fecha de la oferta
            !Nuoritem_Otrab = 0
            !CanPed_Otrab = CAOPE#
            !Especi_oTRAB = ESPECI$
            !preven_Otrab = PREVEOTR#
            !mone_Otrab = MONEOTR%
            !CotiMone_Otrab = COMOTR#
            !LENTREGA_OTRAB = LUENOTR$
            !Ajusta_Otrab = 0
            !VENDED_OTRAB = VEOTR%
            !ComiVend_Otrab = COVEOTR$
            !HsInge_Otrab = HSINGE%
            !HsAvor_Otrab = HSAVOR%
            !HsCCal_Otrab = HSCCAL%
            !HsCald_Otrab = HSCAL%
            !HsMeca_Otrab = HSMEC%
            !HsMoInt_Otrab = HSMOIN%
            !HsMoExt_Otrab = HSMOEXT%
            !HsIngeInf_Otrab = HSINGINF%
            !HsAvorInf_Otrab = HSAVIN%
            !HsCCalInf_Otrab = HSCCINF%
            !HsCaldInf_Otrab = HSCAINF%
            !HsMecaInf_Otrab = HSMEINF%
            !HsMointInf_Otrab = HSMOINF%
            !HsMoExtInf_Otrab = HSMOEINF%
            !Observa_Otrab = OBSOTR$
          .Update
       End With
       End If
       'ahora copio la lista de grupos
       Call ABREORTRA
       'me fijo que no tenga ya grupos cargados
       GrupOtra.FindFirst " nro_otrab = '" & NUOTR2$ & "' "
       If GrupOtra.NoMatch = True Then
        GoTo 17
       Else
       If COMALTER%("La Especificación de Venta " & NUOTR2$ & " \ya Contiene Grupos, ¿Desea Sobre-Escribirlos?\\Si,Sobre-escribir.\\No, Cancelar.") = 1 Then
        Exit Sub
       Else
        GrupOtra.FindFirst " nro_otrab = '" & NUOTR2$ & "' "
        If GrupOtra.NoMatch = False Then
            GrupOtra.Delete
        End If
14      GrupOtra.FindNext " nro_otrab = '" & NUOTR2$ & "' "
        If GrupOtra.NoMatch = True Then
          GoTo 17
        Else
          GrupOtra.Delete
          GoTo 14
        End If
       End If
       End If
       '
17     SQLX$ = " SELECT * FROM GRUPOTRA "
       SQLX$ = SQLX$ + " WHERE Nro_Otrab = '" & NUOTR1$ & "' "
       Set LIGRUTMP = OTrabajo.OpenRecordset(SQLX$, 4)
       '
       With LIGRUTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          GoTo 36
        Else
         Do While .EOF = False
           NUGRUP$ = !Nro_Grup
           PROGRU$ = !PROG_GRUP
           COGRU$ = !COMP_GRUP
           DEGRU$ = !DENO_GRUP
           OBSGRU$ = !OBSERVA_GRUP
        With GrupOtra
          .AddNew
            !Nro_Otrab = NUOTR2$
            !Nro_Grup = NUGRUP$
            !PROG_GRUP = PROGRU$
            !COMP_GRUP = COGRU$
            !DENO_GRUP = DEGRU$
            !OBSERVA_GRUP = OBSGRU$
          .Update
         End With
           .MoveNext
          Loop
          End If
         End With
         '
     'AHORA COPIO LA LISTA DE MATERIALES CORRESPONDIENTES
       SQLX$ = " SELECT * FROM LISTAMAT "
       SQLX$ = SQLX$ + " WHERE Nro_Otrab = '" & NUOTR1$ & "' "
       Set LIMATTMP = OTrabajo.OpenRecordset(SQLX$, 4)
       '
       With LIMATTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          GoTo 36
        Else
         Do While .EOF = False
           NUGRUP$ = !Nro_Grup
           NUOR% = !NuOrden
           CANTI1# = !CANTID1
           DESCRICO$ = !Descri_Cor
           NUPLA$ = !NPlano
           POS$ = !Posic
           MATER$ = !Material
           CODIMA$ = !CODEXTE1
           KIL# = !Kilos
           CLA$ = !CLAVE1
           DESLAR$ = !Descri_Lar
           With ListaMat
             .AddNew
                !Nro_Otrab = NUOTR2$
                !Nro_Grup = NUGRUP$
                !NuOrden = NUOR%
                !CANTID = CANDTI1#
                !Descri_Cor = DESCRICO$
                !NPlano = NUPLA$
                !Posic = POS$
                !Material = MATER$
                !CODEXTE = CODIMA$
                !CLAVE = CLA$
                !Descri_Lar = DESLAR$
             .Update
           End With
           .MoveNext
          Loop
          End If
         End With
     '
36   Call GENEINDIOTRAB
     Text1 = TRIM$(NUOTR2$)
     '
99   Screen.MousePointer = 1
     Command21.Enabled = True
     '
   End If
   If SSTab1.Tab = 2 Then 'copia de grupos
     If DERACCE%("ALTNUEXCO", "Alta de Grupos por Copia") < 2 Then
      Exit Sub
     End If
     '
     Screen.MousePointer = 1
     '
     VDEF$ = Space$(64) + String$(6, 0)
     '
     OBj$ = OBJPLA$("DEACODEV", VDEF$)
     If OBj$ = "" Then GoTo 99
     VDEF$ = OBj$
     '
     NUOTR1$ = TRIM$(Mid$(OBj$, 1, 16))
     NUOTR2$ = TRIM$(Mid$(OBj$, 17, 48))
     '//////////////VALIDACIONES////////////
     If TRIM$(NUOTR1$) = "" Then
        Call COMUNI("Número de Orden de Trabajo no válida")
        GoTo 9
     End If
     '
     If TRIM$(NUOTR2$) = "" Then
        Call COMUNI("Número de Orden de Trabajo no válida")
        GoTo 9
     End If
     '
     Call ABREORTRA
     SQLX$ = " SELECT * FROM GRUPOTRA "
     SQLX$ = SQLX$ + " WHERE Nro_Otrab = '" & NUOTR1$ & "' "
     Set LIGRUTMP = OTrabajo.OpenRecordset(SQLX$, 4)
     '
     With LIGRUTMP
       On Error Resume Next
       .MoveFirst
     If Err Then
        GoTo 36
     Else
     Do While .EOF = False
       NUGRUP$ = !Nro_Grup
       PROGRU$ = !PROG_GRUP
       COGRU$ = !COMP_GRUP
       DEGRU$ = !DENO_GRUP
       OBSGRU$ = !OBSERVA_GRUP
       With GrupOtra
        .AddNew
        !Nro_Otrab = NUOTR2$
        !Nro_Grup = NUGRUP$
        !PROG_GRUP = PROGRU$
        !COMP_GRUP = COGRU$
        !DENO_GRUP = DEGRU$
        !OBSERVA_GRUP = OBSGRU$
        .Update
       End With
     .MoveNext
     Loop
     If COMALTER%("¿Desea Incluir en la Copia las Listas de Materiales Correspondientes?\\No,Incluir.\\Si, Incluir.") = 1 Then
        Exit Sub
     Else
     'AHORA COPIO LA LISTA DE MATERIALES CORRESPONDIENTES
       SQLX$ = " SELECT * FROM LISTAMAT "
       SQLX$ = SQLX$ + " WHERE Nro_Otrab = '" & NUOTR1$ & "' "
       Set LIMATTMP = OTrabajo.OpenRecordset(SQLX$, 4)
       '
       With LIMATTMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          GoTo 36
        Else
         Do While .EOF = False
           NUGRUP$ = !Nro_Grup
           NUOR% = !NuOrden
           CANTI1# = !CANTID1
           DESCRICO$ = !Descri_Cor
           NUPLA$ = !NPlano
           POS$ = !Posic
           MATER$ = !Material
           CODIMA$ = !CODEXTE1
           KIL# = !Kilos
           CLA$ = !CLAVE1
           DESLAR$ = !Descri_Lar
           With ListaMat
             .AddNew
                !Nro_Otrab = NUOTR2$
                !Nro_Grup = NUGRUP$
                !NuOrden = NUOR%
                !CANTID = CANDTI1#
                !Descri_Cor = DESCRICO$
                !NPlano = NUPLA$
                !Posic = POS$
                !Material = MATER$
                !CODEXTE = CODIMA$
                !CLAVE = CLA$
                !Descri_Lar = DESLAR$
             .Update
           End With
           .MoveNext
          Loop
          End If
        End With
     End If
   End If
   End With
   End If
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
        Command22.ToolTipText = "Alta de Nueva Especificación de Venta por Copia de una Existente"
    End If
    If SSTab1.Tab = 2 Then
        Command22.ToolTipText = "Alta de Nueva Lista de Grupos a partir de una Lista Existente"
    End If
    '
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub Command7_Click()
    '
    If TRIM$(Text1) <> "" Then
        '
        FORIPRE$ = CONTROL$("", "LISTGRUP")
        '
        If TRIM$(FORIPRE$) <> "" Then
          '
          Screen.MousePointer = 11
          FECHDOC$ = Text9
          FECHFIN$ = Text10
          NUMEDOC$ = TRIM$(Text1)
          DENOOBRA$ = Text3
          DESTIDOC% = Val(Label22) 'CLIENTE
          RPLA$ = TRIM$(Text15)
          '
          CODPARAM$(1) = "RASO-CLI" 'CLIENTE
          CODPARAM$(2) = "REF-MAT1" 'DENOMINACION DE LA ESP. DE VENTA
          CODPARAM$(3) = "NUME-COM" 'NUMERO DE ESPECIFICACION DE VENTA
          CODPARAM$(4) = "FECH-EMI" 'FECHA EMISION DEL LISTADO DE GRUPOS
          CODPARAM$(5) = "FECH-VEN" 'FECHA DE ENTREGA
          CODPARAM$(6) = "HOJA-NUM" 'NUMERO DE HOJA
          CODPARAM$(7) = "HOJA-TOT" 'TOTAL DE HOJAS
          CODPARAM$(8) = "R-PLANO"  'NUMERO DE REVISION
          CAPARDOC% = 8
          
          DOCPARAM(1) = RASOCLI(DESTIDOC%)
          DOCPARAM(2) = DENOOBRA$
          DOCPARAM(3) = NUMEDOC$
          DOCPARAM(4) = FECHDOC$
          DOCPARAM(5) = FECHFIN$
          DOCPARAM(6) = ""
          DOCPARAM(7) = ""
          DOCPARAM(8) = RPLA$
          '
          CAITAB1% = 0
          CODTABU1$(1) = "COD-MAT" ' CODIGO DE GRUPO
          CODTABU1$(2) = "DES-MAT" ' DENOMINACION DEL GRUPO
          CODTABU1$(3) = "MEDIDA-1" ' PROGRAMA
          CODTABU1$(4) = "MEDIDA-2" '
          CODTABU1$(5) = "MEDIDA-3" ' OBSERVACIONES
          CACOLTAB1% = 5
          '
          For U& = 1 To ULTREG&("!GRUORTRA")
            Z$ = REGLEIDO$("!GRUORTRA", U&)
            NROGRU$ = Mid$(Z$, 1, 8)
            PROGRU$ = Mid$(Z$, 9, 12)
            COMGRU$ = Mid$(Z$, 21, 12)
            DENGRU$ = Mid$(Z$, 33, 64)
            OBSGRU$ = Mid$(Z$, 97, 64)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = NROGRU$
            TETABU1$(2, CAITAB1%) = DENGRU$
            TETABU1$(3, CAITAB1%) = PROGRU$
            TETABU1$(4, CAITAB1%) = COMGRU$
            TETABU1$(5, CAITAB1%) = OBSGRU$
          Next U&
          '
          NOCONFIRM% = 1
          Call PRINTDOC("LISTGRUP")                   ' LISTADO DE GRUPOS
          '
        End If
199     Call CIERRARCH("*.*")
        Call BAJALDISCO
        Screen.MousePointer = 1
        '
        Call CIERRARCH("*.*")
        '
        Exit Sub
        '
999     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        '
    Else
        Call COMUNI("Debe Seleccionar una Especificación de Venta.")
    End If
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(RASOCLI$(NCLIE%), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE%)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
    Call Command14_Click
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   openforms = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   If TRIM(Text1) <> "" Then
    Call SELECHO("VENDEDOR")
    VDEVU$ = VALACT1$("VENDEDOR")
    If TRIM$(VDEVU$) <> "" Then
         COTEXT(5).TEXT = VDEVU$
    End If
   Else
    Call COMUNI("Debe Seleccionar una Especificación de Venta.")
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   'Call GRAGENCLI
   'Call GRANOTACLI
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPCESPVTA04.Show
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPORTRA04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    If TRIM$(Text1.TEXT) <> "" Then
        Call GRABOT
        Call SUMAHS
    Else
        Call COMUNI("Debe Seleccionar una Especificación de Venta.")
    End If
End Sub
Sub SUMAHS()
    XX1 = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
    Text34 = FORMATNUM(XX1, "F6.2")
End Sub
Private Sub Command17_Click()
   'COMIE = Text28.SelStart
   'LOSEL = Text28.SelLength
   'Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   'Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
   Call CompactDatabase(LUDAT$ + "\BDACCESS\OTRABAJO.MDB", LUDAT$ + "\BDACCESS\OTRABAJNUE.MDB")
End Sub

Private Sub Command2_Click()
   'Call GRAGENCLI
   'Call GRANOTACLI
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   'Call CARDATCLIE
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call ABRECLIEN
     On Error Resume Next
     With Deudor12
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             VATEX$ = .Fields(KKI% - 1).Value
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   '
   If DERACCE%("ALTACLI", "Alta de Nueva Orden de Trabajo") < 2 Then
      Exit Sub
   End If
   '
9  Screen.MousePointer = 1
   '
   'FALTA PONER EL CODIGO PARA QUE TRAIGA LAS FECHAS POR DEFAULT
   '
   VDEF$ = Space$(64) + String$(6, 0)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 67, 2)
   '
   OBj$ = OBJPLA$("NUEOTRA", VDEF$)
   If OBj$ = "" Then GoTo 99
   VDEF$ = OBj$
   '
   NUOTR$ = TRIM$(Mid$(OBj$, 1, 16))
   DESCRIOTR$ = TRIM$(Mid$(OBj$, 17, 48))
   NUCLIOTR% = CVI(Mid$(OBj$, 65, 2))
   RASOCLIOTR$ = RASOCLI$(NUCLIOTR%)
   FECOTR% = CVI(Mid$(OBj$, 67, 2))
   FEENOTR% = CVI(Mid$(OBj$, 69, 2))
   
   '//////////////VALIDACIONES////////////
   If TRIM$(NUOTR$) = "" Then
      Call COMUNI("Número de Orden de Trabajo no válida")
      GoTo 9
   End If
   '
   Call ABREORTRA
   OrdeTrab.FindFirst "Nro_Otrab = '" & NUOTR$ & "' "
   If OrdeTrab.NoMatch = False Then
      Call COMUNI("Numero de Orden de Trabajo no Válido o Pre-Existente")
      GoTo 9
   End If
   '
   If TRIM$(DESCRIOTR$) = "" Then
      Call COMUNI("Debe Ingresar una Descripcion para la Orden de Trabajo")
      GoTo 9
   End If
   '
   If NUCLIOTR% <= 0 Then
      Call COMUNI("Cliente no Valido.")
      GoTo 9
   End If
   '
   If FEENOTR% < FECOTR% Then
      Call COMUNI("La Fecha de Entrega Debe ser Superior o Igual a la Fecha de Generación")
      GoTo 9
   End If
   '
   jjj = CVINT(Date)
   If FECOTR > jjj Then
      Call COMUNI("La Fecha de Generación no puede ser Superior al dia de la Fecha.")
      GoTo 9
   End If
   '
   Call ABREORTRA
   With OrdeTrab
      .AddNew
        !Nro_Otrab = NUOTR$
        !Refe_Otrab = DESCRIOTR$
        !Nclie_Otrab = NUCLIOTR%
        !FEMI_oTRAB = CVDAT(FECOTR%)
        !Fentre_Otrab = CVDAT(FEENOTR%)
        !uSUCONF_OTRAB = USERNAME 'usuario de confección
        !UsuCont_Otrab = "" 'usuario que controlo
        !UsuApro_Otrab = "" 'usuario que aprobara
        !NroOcom_Otrab = "" 'numero de orden de compra del cliente
        !FeOcom_Otrab = Date 'fecha de pedido del cliente
        !NroOfer_Otrab = "" 'Numero de Oferta
        !FeOfer_Otrab = Date 'fecha de la oferta
        !Nuoritem_Otrab = 0
        !CanPed_Otrab = 0
        !Especi_oTRAB = ""
        !preven_Otrab = 0
        !mone_Otrab = 5
        !CotiMone_Otrab = 0
        !LENTREGA_OTRAB = " "
        !Ajusta_Otrab = 0
        !VENDED_OTRAB = 0
        !ComiVend_Otrab = 0
        !HsInge_Otrab = 0
        !HsAvor_Otrab = 0
        !HsCCal_Otrab = 0
        !HsCald_Otrab = 0
        !HsMeca_Otrab = 0
        !HsMoInt_Otrab = 0
        !HsMoExt_Otrab = 0
        !HsIngeInf_Otrab = 0
        !HsAvorInf_Otrab = 0
        !HsCCalInf_Otrab = 0
        !HsCaldInf_Otrab = 0
        !HsMecaInf_Otrab = 0
        !HsMointInf_Otrab = 0
        !HsMoExtInf_Otrab = 0
        !Observa_Otrab = " "
      .Update
   End With
   '
   Call GENEINDIOTRAB
   Text1 = TRIM$(NUOTR$)
   '
99 Screen.MousePointer = 1
   Command21.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   SSTab1.Tab = 0
   Text20.TEXT = " Suspendido"
   Text20.Refresh
End Sub

Private Sub Command25_Click()
   If Text1 <> "" Then
        Call PRIORVEN
   Else
        Call COMUNI("Debe Seleccionar una Especificación de Venta.")
   End If
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub

Private Sub Command5_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text24.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text24.TEXT)
 End If
End Sub

Private Sub Command6_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text12.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text12.TEXT)
 End If
End Sub

Private Sub Command8_Click()
   'Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Call COMUNI("No Disponible en la presente versión")
End Sub

Private Sub Command9_Click()
   '
   Call GRABOT
   Call SELECHO("INDIOTRA")
   Nuotra$ = TRIM$(VALACT1$("INDIOTRA"))
   Text1.TEXT = TRIM$(Nuotra$)
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Frame3_DBLCLICK()
    Call List1_DblClick
End Sub

Private Sub Label2_dblClick()
    '
    'TAMONED VA A ESTAR DEFINIDA COMO:
    '1=PESOS
    '2=DÓLARES
    '3=EUROS
    '4=REALES
     Select Case UCase(Label2)
       Case "PESOS"
        Label2.ForeColor = RGB(255, 0, 0)
        Label2 = "Dólares"
        MONEMI% = 2
        Label2.ToolTipText = "Doble-Click para 'Euros'"
        
      Case "DÓLARES"
        Label2.ForeColor = vbYellow
        Label2 = "Euros"
        MONEMI% = 3
        Label2.ToolTipText = "Doble-Click para 'Reales'"
        
      Case "EUROS"
        Label2.ForeColor = vbMagenta
        Label2 = "Reales"
        Label2.ToolTipText = "Doble-Click para 'Pesos'"
        
        MONEMI% = 4
      Case "REALES"
        Label2.ForeColor = RGB(0, 0, 255)
        Label2.Caption = "Pesos"
        Label2.ToolTipText = "Doble-Click para 'Dólares'"
        MONEMI% = 1
     End Select
     '
End Sub

Private Sub Label25_Change()
   Label35 = Label25
   SUMA = Val(Label25) + Val(Label29)
   Label44 = FORMATNUM(SUMA, "F6.1")
End Sub

Private Sub Label29_Change()
   Label40 = Label29
   SUMA = Val(Label25) + Val(Label29)
   Label44 = FORMATNUM(SUMA, "F6.1")
End Sub

Private Sub Label37_DblClick()
   ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left))
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + SSTab1.Top + Frame18.Top))
   VTB% = VENTABU%("GRUORTRA/" + ATRI$)
   '
   If VTB% >= 0 Then
       List3.Clear
       For U& = 1 To ULTREG&("!GRUORTRA")
         Z$ = REGLEIDO$("!GRUORTRA", U&)
         NROGRU$ = Mid$(Z$, 1, 8)
         PROGRU$ = Mid$(Z$, 9, 12)
         COMGRU$ = Mid$(Z$, 21, 12)
         DENGRU$ = Mid$(Z$, 33, 64)
         OBSGRU$ = Mid$(Z$, 97, 64)
         '
         Y$ = Space$(128)
         Call REPLA(Y$, NROGRU$, 1, 8)
         Call REPLA(Y$, PROGRU$, 10, 6)
         Call REPLA(Y$, COMGRU$, 17, 6)
         Call REPLA(Y$, DENGRU$, 24, 48)
         Call REPLA(Y$, OBSGRU$, 73, 24)
         List3.AddItem Y$
       Next U&
       MODIFIC% = 1
     Else
       REBLA$ = REGBLAN$("GRUORTRA")
       JJ& = 0
       For U& = 1 To List3.ListCount
         Y$ = List3.List(U& - 1)
         NROGRU$ = Mid$(Y$, 1, 8)
         PROGRU$ = Mid$(Y$, 10, 6)
         COMGRU$ = Mid$(Y$, 17, 6)
         DENGRU$ = Mid$(Y$, 24, 48)
         OBSGRU$ = Mid$(Y$, 73, 24)
         '
         Z$ = REBLA$
         Call REPLA(Z$, NROGRU$, 1, 8)
         Call REPLA(Z$, PROGRU$, 9, 12)
         Call REPLA(Z$, COMGRU$, 21, 12)
         Call REPLA(Z$, DENGRU$, 33, 64)
         Call REPLA(Z$, OBSGRU$, 97, 64)
         JJ& = JJ& + 1
         Call GRAREG("!GRUORTRA", Z$, JJ&)
       Next U&
       Call SETULTREG("!GRUORTRA", JJ&)
   End If
   '
End Sub

Private Sub Label5_Click()
    Call List1_DblClick
End Sub

Private Sub LiGeClaCli_Click()
    Call Command13_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub List1_DblClick()
        'GENERO EL DAT
        If TRIM$(Text1.TEXT) <> "" Then
        NUOTRAB$ = TRIM$(Text1.TEXT)
        Call ABREORTRA
        SQLX$ = "SELECT * FROM PLACOB "
        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
        J& = 0
        With PLACOTMP
          '
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do Until .EOF = True
              Y$ = REGBLAN("!CAPLACO")
              '
              NUTRA$ = TRIM$(!Nro_Otrab)
              Call REPLA(Y$, NUTRA$, 1, 16)
              NUORITOT% = !NuOrItem
              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
              DESCRIOT$ = TRIM$(!Refe_Pcob)
              Call REPLA(Y$, DESCRIOT$, 19, 64)
              IMPOT# = !Impo_Pcob
              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
              FECOT% = CVINT(!FeVen_Pcob)
              Call REPLA(Y$, MKI(FECOT%), 91, 2)
              If IsNull(!NroFac_Pcob) = True Then
                NROFAOT$ = " "
              Else
                NROFAOT$ = TRIM$(!NroFac_Pcob)
              End If

              Call REPLA(Y$, NROFAOT$, 93, 16)
              MONEOT% = !Moneda
              Call REPLA(Y$, MKI(MONEOT%), 109, 2)
              COTIOT# = !Coti_Pcob
              Call REPLA(Y$, MKD(COTIOT#), 111, 8)
              '
              J& = J& + 1
              Call GRAREG("!CAPLACO", Y$, J&)
            .MoveNext
            Loop
          End If
          On Error GoTo 0
        End With
        '
        Call SETULTREG("!CAPLACO", J&)
        Call CIERRARCH("!CAPLACO")
        '
        'Call GENEPLACO 'GENERO EL DAT
        PRF$ = "CAPLACO"
        ATRI$ = "/NC"
        ATRI$ = ATRI$ + "/TITUPAN=Plan de Cobranza"
        VTB% = VENTABU%(PRF$ + ATRI$)
        If VTB% >= 0 Then
          'GUARDO O EDITO LOS REGISTROS EN TABLA
          'ACA VA RUTINA DE GRABACION
          Call ABREORTRA
          J& = ULTREG("!CAPLACO")
          For KKIL& = 1 To J&
             XX$ = REGLEIDO$("!CAPLACO", KKIL&)
             NUOTRAB$ = TRIM$(Text1.TEXT)
             ITE% = CVI(Mid$(XX$, 17, 2))
             'VALIDACION SI ITEM VALE 0
             If ITE% <= 0 Then GoTo 7
             '
             DESCRI$ = Mid$(XX$, 19, 64)
             IMPOT# = CVD(Mid$(XX$, 83, 8))
             FECOTR% = CVI(Mid$(XX$, 91, 2))
             NUFAC = Mid$(XX$, 93, 16)
             With PlanCob
             '/// Aca FindFirst de Item+NUOTRAB$, Si existe
             ' se edita el registro, sino,
             'se crea registro nuevo////
             SQLX$ = " Nro_Otrab = '" & NUOTRAB$ & "' "
             SQLX$ = SQLX$ + " AND NuOrItem = " & ITE% & " "
             .FindFirst SQLX$
             If .NoMatch = True Then
               .AddNew
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
               .Update
              Else
               .Edit
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
               .Update
              End If
              End With
7         Next KKIL&
        End If
    End If
'    '/////////*******ACA GENERO NUEVAMENTE EL DAT*******///////////
'        NUOTRAB$ = TRIM$(Text1.TEXT)
'        Call ABREORTRA
'        SQLX$ = "SELECT * FROM PLACOB "
'        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
'        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
'        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
'        J& = 0
'        With PLACOTMP
'          '
'          On Error Resume Next
'          .MoveFirst
'          xxsd = .RecordCount
'          MsgBox xxsd
'          If Err < 1 Then
'            On Error GoTo 0
'            Do Until .EOF = True
'              Y$ = REGBLAN("!CAPLACO")
'              '
'              NUTRA$ = TRIM$(!Nro_Otrab)
'              Call REPLA(Y$, NUTRA$, 1, 16)
'              NUORITOT% = !NuOrItem
'              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
'              DESCRIOT$ = TRIM$(!Refe_Pcob)
'              Call REPLA(Y$, DESCRIOT$, 19, 64)
'              IMPOT# = !Impo_Pcob
'              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
'              FECOT% = CVINT(!FeVen_Pcob)
'              Call REPLA(Y$, MKI(FECOT%), 91, 2)
'              If IsNull(!NroFac_Pcob) = True Then
'                NROFAOT$ = " "
'              Else
'                NROFAOT$ = TRIM$(!NroFac_Pcob)
'              End If
'
'              Call REPLA(Y$, NROFAOT$, 93, 16)
'              MONEOT% = !Moneda
'              Call REPLA(Y$, MKI(MONEOT%), 109, 2)
'              COTIOT# = !Coti_Pcob
'              Call REPLA(Y$, MKD(COTIOT#), 111, 8)
'              '
'              J& = J& + 1
'              Call GRAREG("!CAPLACO", Y$, J&)
'            .MoveNext
'            Loop
'          End If
'          On Error GoTo 0
'        End With
'        '
'        Call SETULTREG("!CAPLACO", J&)
'        Call CIERRARCH("!CAPLACO")
        '
    '/////////******ACA GUARDO EL PLAN EDITADO******///////////////
     With PlanCob
20   .FindFirst "Nro_Otrab = '" & NUOTRAB$ & "' "
     If .NoMatch = False Then
       .Delete
       GoTo 20
     End If
     '
     For U& = 1 To ULTREG("!CAPLACO")
        Y$ = REGLEIDO$("!CAPLACO", U&)
        'NUOTR$ = Mid$(Y$, 1, 16)
        ITE1% = CVI(Mid$(Y$, 17, 2))
        REFE1$ = Mid$(Y$, 19, 64)
        IMP1# = CVD(Mid$(Y$, 83, 8))
        FECVEN% = CVI(Mid$(Y$, 91, 2))
        NFAC1$ = Mid$(Y$, 93, 16)
        MON1% = CVI(Mid$(Y$, 109, 2))
        COTMON1# = CVD(Mid$(Y$, 111, 8))
       '
       .AddNew
        !Nro_Otrab = NUOTRAB$
        !NuOrItem = ITE1%
        !Refe_Pcob = REFE1$
        !Impo_Pcob = IMP1#
        !FeVen_Pcob = CVDAT(FECVEN%)
        !NroFac_Pcob = NFAC1$
        !Moneda = MON1%
        !Coti_Pcob = COTMON1#
       .Update
     Next U&
   End With
   '//////*********/////////////
    List1.Clear
    Call CARGOPLACO
    '
End Sub


Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuOrItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !Moneda, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
      
    
End Sub



Private Sub NueOT_Click()
    Call Command21_Click
End Sub

Private Sub Text10_Change()
   MODIFIC% = 1
End Sub

Private Sub Text11_GotFocus()
    Text25.SetFocus
End Sub

Sub GRABOT() 'GRABAR ORDEN DE TRABAJO
   '
   If YACAR% < 1 Or MODIFIC% < 1 Then
     Exit Sub
   End If
   '
   If DERACCE%("GRADATCLI", "Grabar Datos de Ordenes de Trabajo") < 2 Then
       Exit Sub
   End If
   'VALIDACIONES
   '
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Orden de Trabajo no válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("INDIOTRA", Text1) = "" Then
      Call COMUNI("Número de Orden de Trabajo no válido.")
      Exit Sub
   End If
   '
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Descripción de Orden de Trabajo.")
      Exit Sub
   End If
   '
   If FECHANUM(Text10) < FECHANUM(Text9) Then
       Call COMUNI("Fecha de Entrega no válida.")
       Exit Sub
   End If
   '
   If TRIM$(Text17) = "" Then
     Call COMUNI("Falta Especificación de Venta.")
     Exit Sub
   End If
   '
   If Val(TRIM$(Text36)) < 1 Then
       Call COMUNI("Cantidad no Válida.")
       Exit Sub
   End If
   '
   If CDbl(Val(TRIM$(Text38))) <= 0.005 Then
       Call COMUNI("Importe de Venta no válido.")
       Exit Sub
   End If
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM ORDETRAB"
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 2)
   With OTTMP
     .FindFirst "  NRO_oTRAB= '" & NUOTR$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        'Call COMUNI("No Existe la Orden de Trabajo")
        Exit Sub
     End If
     On Error GoTo 0
     '
     .Edit
        !Especi_oTRAB = Text17
        !Fentre_Otrab = CVDAT(CVINT(Text10))
        !Refe_Otrab = TRIM$(Text3)
        !NroOcom_Otrab = TRIM$(Text25)
        !FeOcom_Otrab = CVDAT(CVINT(Text24))
        !NroOfer_Otrab = TRIM$(Text13)
        !FeOfer_Otrab = CVDAT(CVINT(Text12))
        !CanPed_Otrab = Val(Text36)
        !preven_Otrab = Val(Text38)
        !LENTREGA_OTRAB = Text2
        'GRABO HORAS
        If Val(Text30) > 0 Then
           !HsCald_Otrab = Val(Text30)
        End If
        If Val(Text29) > 0 Then
           !HsMeca_Otrab = Val(Text29)
        End If
        If Val(Text28) > 0 Then
           !HsMoInt_Otrab = Val(Text28)
        End If
        If Val(Text32) > 0 Then
           !HsMoExt_Otrab = Val(Text32)
        End If
        If Val(Text19) > 0 Then
            !HsInge_Otrab = Val(Text19)
        End If
        If Val(Text4) > 0 Then
            !HsAvor_Otrab = Val(Text4)
        End If
        If Val(Text14) > 0 Then
            !HsCCal_Otrab = Val(Text14)
        End If
        '
        If Val(COTEXT(5)) > 0 Then
            !VENDED_OTRAB = Val(COTEXT(5))
        End If
        !ComiVend_Otrab = Val(Text21)
        'para grabar la moneda:
        'TAMONED VA A ESTAR DEFINIDA COMO:
        '1=PESOS
        '2=DÓLARES
        '3=EUROS
        '4=REALES
        Select Case UCase(Label2)
         Case "PESOS"
           MONEOT% = 1
         Case "DÓLARES"
           MONEOT% = 2
         Case "EUROS"
           MONEOT% = 3
         Case "REALES"
           MONEOT% = 4
        End Select
        !mone_Otrab = MONEOT%
        !Observa_Otrab = Text35
        If TRIM$(Text6) = "" Then
            !FeRev_Otrab = Date
        Else
            !FeRev_Otrab = Text6
        End If
        !NuRev_Otrab = Text15
        '
     .Update
   End With
   '
   With GrupOtra
20   .FindFirst "Nro_Otrab= '" & NUOTR$ & "' "
     If .NoMatch = False Then
       .Delete
       GoTo 20
     End If
     '
     For U& = 1 To List3.ListCount
       Y$ = List3.List(U& - 1)
       NROGRU$ = Mid$(Y$, 1, 8)
       PROGRU$ = Mid$(Y$, 10, 6)
       COMGRU$ = Mid$(Y$, 17, 6)
       DENGRU$ = Mid$(Y$, 24, 48)
       OBSGRU$ = Mid$(Y$, 73, 24)
       '
       .AddNew
       !Nro_Otrab = NUOTR$
       !Nro_Grup = NROGRU$
       !PROG_GRUP = PROGRU$
       !COMP_GRUP = COMGRU$
       !DENO_GRUP = DENGRU$
       !OBSERVA_GRUP = OBSGRU$
       .Update
     Next U&
   End With
   '
   MODIFIC% = 0
   Call GENEINDIOTRAB
   '
End Sub

Private Sub Text12_Change()
    MODIFIC% = 1
End Sub

Private Sub Text13_Change()
    MODIFIC% = 1
End Sub

Private Sub Text14_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) + Val(Text4) + Val(Text14)
   Label25 = FORMATNUM(SUMA, "F6.1")
End Sub

Private Sub Text14_LOSTFOCUS()
    Text14 = FORMATNUM(Val(Text14), "F6.1")
End Sub

Private Sub Text15_Change()
    MODIFIC = 1
End Sub

Private Sub Text17_Change()
    MODIFIC% = 1
End Sub

Private Sub Text18_GotFocus()
    Text25.SetFocus
End Sub

Private Sub Text2_Change()
'   Text4 = ""
'   If Val(TRIM$(Text1)) > 0 Then
'     NCLIE% = Val(TRIM$(Text2))
'     If NCLIE% > 0 Then
'       Text4 = RASOCLI(NCLIE%)
'     End If
'   End If
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
End Sub

Private Sub TEXT21_LOSTFOCUS()
    '
    Text21 = FORMATNUM$(Val(TRIM(Text21)), "F6.2")
    '
End Sub

Private Sub Text35_Change()
   MODIFIC% = 1
End Sub

Private Sub Text38_lostfocus()
    '
    MODIFIC% = 1
    Text38 = FORMATNUM$(Val(Text38), "F12.2")
    '
End Sub

Private Sub Text4_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) + Val(Text4) + Val(Text14)
   Label25 = FORMATNUM(SUMA, "F6.1")
End Sub

Private Sub Text4_LOSTFOCUS()
    Text4 = FORMATNUM(Val(Text4), "F6.1")
End Sub

Private Sub Text5_GotFocus()
    Text25.SetFocus
End Sub

Private Sub COTEXT_Change(Index As Integer)
ECO$ = ""
If TRIM$(COTEXT(5).TEXT) <> "" Then
    TV11% = TIPFIELD("VENDEDOR", 1)
    LF11% = LENFIELD("VENDEDOR", 1)
    ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
    ECO$ = TRIM$(ECOARCH$("VENDEDOR", ARGU$))
End If
If ECO$ <> "" Then
   DETEXT(5).Caption = ECO$
Else
   DETEXT(5).Caption = "No Definido"
End If
'   '
'   ECO$ = ""
'   If TRIM$(COTEXT(Index).TEXT) <> "" Then
'      TV11% = TIPFIELD(ARCHEX$(Index), 1)
'      LF11% = LENFIELD(ARCHEX$(Index), 1)
'      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
'      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
'   End If
'   '
'   If ECO$ <> "" Then
'        DETEXT(Index).Caption = ECO$
'      Else
'        DETEXT(Index).Caption = "No Definido"
'   End If
'   '
   MODIFIC% = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text36.TEXT = ""
   Text38.TEXT = ""
   
   'Text26.TEXT = ""
'   Text27.TEXT = ""
'   Text20.TEXT = ""
  Text4.TEXT = ""
  Text14.TEXT = ""
  Text19.TEXT = ""
   Text28.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
'   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
'   Text33.TEXT = ""
'   Text34.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   '
'   For KK% = 0 To 7
'     COTEXT(KK%).TEXT = ""
'   Next KK%
   '
'   Text28.TEXT = ""
   '
     Text23 = ""
     Text35 = ""
     Text36 = ""
     Text37 = ""
     'Text1 = ""
     Text3 = ""
     '
     Text18 = ""
     Text5 = ""
     Text7 = ""
     Text8 = ""
     '
     Text9 = ""
     Text10 = ""
     Text25 = ""
     Text24 = ""
     Text11 = ""
     Text13 = ""
     Text12 = ""
     Text17 = ""
'   List2.Clear
   List3.Clear
   List1.Clear
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
'   Call GRAGENCLI
'   Call GRANOTACLI
'   Call CARDATCLIE
'
  Call CIERRARCH("*.*")
'   If AGRECLI% > 0 Then
'      Call SORTCLI(1)
'      AGRECLI% = 0
'   End If
  
  Call BAJALDISCO
  Call FINAL("")
  ' Hide
'   If APLINVO$ <> "" Then
'       Call FINAL("")
'     Else
'       Hide
'   End If
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
    Call COMMAND15_CLICK
End Sub

Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub List_Click()
   Call Command13_Click
End Sub

Private Sub List3_DblClick()
   '
   LISMAT04.Label31 = Text1
   LISMAT04.Label7 = Text3
   LISMAT04.Label8 = Mid$(List3.TEXT, 24, 48)
   LISMAT04.Label2 = Left$(List3.TEXT, 8)
   LISMAT04.Label5 = Text18
   LISMAT04.Text9 = Text9
   LISMAT04.Text10 = Text10
   LISMAT04.Text11 = Text11
   '
   NUOTR$ = TRIM$(LISMAT04.Label31)
   NUGRU$ = TRIM$(LISMAT04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call LISMAT04.CARGOLISMAT
   '
   LISMAT04.Show 1
   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub NueCli_Click()
   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call COMMAND15_CLICK
End Sub

Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub Susp_Cli_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If TRIM$(Text1.TEXT) <> "" Then
      NUOTR$ = TRIM$(Text1.TEXT)
      Call CARGOTR
      Call CARGOLISGRU
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRAGENCLI
   'Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) + Val(Text4) + Val(Text14)
   Label25 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub text19_lostfocus()
   Text19 = FORMATNUM(Val(Text19), "F6.1")
End Sub
Private Sub Text20_Change()
   MODIFIC% = 1
'   If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
'       LTEXT36.Caption = Text20.TEXT
'     Else
'       LTEXT36.Caption = CSTRING$(MKS$(Val(Text20.TEXT)), 4, 11, 2, 2)
'   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 12, 2, 2)
  End If
End Sub

Private Sub Text21_Change()
   MODIFIC% = 1
End Sub



Private Sub Text23_Change()
   MODIFIC% = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC% = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC% = 1
End Sub

Private Sub Text26_Change()
   MODIFIC% = 1
End Sub

Private Sub Text27_Change()
   MODIFIC% = 1
End Sub

Private Sub Text28_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT28_LOSTFOCUS()
    Text28 = FORMATNUM(Val(Text28), "F6.1")
End Sub
Private Sub Text29_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT29_LOSTFOCUS()
    Text29 = FORMATNUM(Val(Text29), "F6.1")
End Sub
Private Sub Text3_Change()
   MODIFIC% = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT30_LOSTFOCUS()
    Text30 = FORMATNUM(Val(Text30), "F6.1")
End Sub
Private Sub Text31_Change()
   MODIFIC% = 1
End Sub

Private Sub Text32_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT32_LOSTFOCUS()
    Text32 = FORMATNUM(Val(Text32), "F6.1")
End Sub
Private Sub Text33_Change()
   MODIFIC% = 1
End Sub

Private Sub Text34_Change()
   MODIFIC% = 1
End Sub
Sub Text34_LOSTFOCUS()
    Text34 = FORMATNUM(Val(Text34), "F6.1")
End Sub
Private Sub LText35_DblClick()
  Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub


Private Sub Text36_Change()
   MODIFIC% = 1
End Sub
Sub TEXT36_LOSTFOCUS()
    Text36 = FORMATNUM(Val(Text36), "F6.1")
End Sub

Private Sub Text6_Change()
    MODIFIC = 1
End Sub

Private Sub Text7_GotFocus()
    Text25.SetFocus
End Sub

Private Sub Text8_GotFocus()
    Text25.SetFocus
End Sub

Private Sub Text9_GotFocus()
    Text25.SetFocus
End Sub

Private Sub VenCli_Click()
With OPMASCL04
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Sub GENECOCLI()
   If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABRECLIEN
      Set CLIENTECO = Badaclie.OpenRecordset("SELECT Codi_Cli, Nume_Cli,Raso_Cli FROM Deudor12 WHERE Stat_Cli > 0 ORDER BY Raso_Cli ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      CLIENTECO.MoveLast
      If Err Then Exit Sub
      '
      CAREC% = CLIENTECO.RecordCount
      If CAREC% < 1 Then Exit Sub
      '
      FIN& = CAREC%
      REBLA$ = REGBLAN$("DEUDOR1")
      REBLB$ = REGBLAN$("INDICLIE")
      JJ& = 0
      With CLIENTECO
        .MoveFirst
        '
10      Call TRACE(20, JJ&, FIN&)
        EL1$ = REBLA$
        Call REPLA(EL1$, MKI$(!Nume_Cli), 1, 2)
        Call REPLA(EL1$, !Raso_Cli, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("DEUDOR1", EL1$, JJ&)
        '
        EL2$ = REBLB$
        Call REPLA(EL2$, !Codi_Cli, 1, 12)
        Call REPLA(EL2$, !Raso_Cli, 13, 50)
        Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
        Call GRAREG("INDICLIE", EL2$, JJ&)
        '
        If JJ& < CAREC% Then
          CLIENTECO.MoveNext
          GoTo 10
        End If
      End With
      '
      CLIENTECO.Close
      Call SETULTREG("DEUDOR1", JJ&)
      Call CIERRARCH("DEUDOR1")
      Call FILESORT("INDICLIE", "", 1, 1, "ASC")
      Call FRESHECHO("DEUDOR1")
      Call FRESHECHO("INDICLIE")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
   End If
End Sub
'
Sub GENDAPIBCLI()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("DEUDOR1")
    REBLA$ = REGBLAN$("DAPIBCLI")
    Call REPLA(REBLA$, USERNAME$, 41, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
    Call REPLA(REBLA$, MKI$(HORANUM%(Time$)), 63, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        PRC$ = UCase$(PROVCLI$(NCLIX%))
        RECORD$ = REBLA$
        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
        If PRC$ = "B" Then
            Call REPLA(RECORD$, Chr$(1), 3, 1)
            Call REPLA(RECORD$, MKS$(100), 5, 4)
          Else
            Call REPLA(RECORD$, Chr$(0), 3, 1)
            Call REPLA(RECORD$, MKS$(0), 5, 4)
            Call REPLA(RECORD$, "EXTRAZONA", 9, 32)
        End If
        Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIX%), RECORD$)
      End If
    Next U&
    Call CIERRARCH("DAPIBCLI")
    Screen.MousePointer = 1
    '
End Sub

Sub CARGOTR() 'CARGA LA FICHA DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM ORDETRAB"
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With OTTMP
     .FindFirst "  NRO_oTRAB= '" & NUOTR$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        'Call COMUNI("No Existe la Orden de Trabajo")
        Exit Sub
     End If
     On Error GoTo 0
     '
     Text3 = !Refe_Otrab
     '
     NCLIE% = !Nclie_Otrab
     Label22 = NCLIE%
     If NCLIE% > 0 Then
       Text18 = RASOCLI(NCLIE%)
       Text5 = DOMICLI$(NCLIE%)
       Text7 = CPOSCLI$(NCLIE%)
       Text8 = LOCACLI$(NCLIE%)
     End If
     '
     Text9 = FECHATEX(CVINT(!FEMI_oTRAB))
     Text10 = FECHATEX(CVINT(!Fentre_Otrab))
     Text25 = !NroOcom_Otrab
     Text24 = FECHATEX(CVINT(!FeOcom_Otrab))
     Text11 = !uSUCONF_OTRAB
     Text13 = !NroOfer_Otrab
     Text12 = FECHATEX(CVINT(!FeOfer_Otrab))
     Text17 = !Especi_oTRAB
     Text36 = !CanPed_Otrab
     Text38 = !preven_Otrab
     Text38 = FORMATNUM$(Val(Text38), "F12.2")
     If IsNull(!FeRev_Otrab) Then
        Text6 = ""
     Else
        Text6 = FECHATEX(CVINT(!FeRev_Otrab))
     End If
     If IsNull(!NuRev_Otrab) Then
        Text15 = ""
     Else
        Text15 = !NuRev_Otrab
     End If
     '
     Text2 = !LENTREGA_OTRAB
     COTEXT(5) = !VENDED_OTRAB
     COMIVE = Val(!ComiVend_Otrab)
     Text21 = FORMATNUM(COMIVE, "F6.2")
     '
     If !mone_Otrab > 0 Then
      'CARGO LA MONEDA
      'TAMONED VA A ESTAR DEFINIDA COMO:
      '1=PESOS
      '2=DÓLARES
      '3=EUROS
      '4=REALES
      Select Case !mone_Otrab
       Case 1
         Label2 = "Pesos"
         Label2.ForeColor = vbBlue
       Case 2
         Label2 = "Dólares"
         Label2.ForeColor = vbRed
       Case 3
         Label2 = "Euros"
         Label2.ForeColor = vbYellow
       Case 4
         Label2 = "Reales"
         Label2.ForeColor = vbMagenta
      End Select
      Label2.FontUnderline = True
     End If
     'CARGO HS
     XX1 = !HsCald_Otrab
     If XX1 > 0.005 Then
        Text30 = FORMATNUM(XX1, "F6.1")
     End If
     XX2 = !HsMeca_Otrab
     If XX2 > 0.005 Then
        Text29 = FORMATNUM(XX2, "F6.1")
     End If
     XX3 = !HsMoInt_Otrab
     If XX3 > 0.005 Then
        Text28 = FORMATNUM(XX3, "F6.1")
     End If
     XX4 = !HsMoExt_Otrab
     If XX4 > 0.005 Then
        Text32 = FORMATNUM(XX4, "F6.1")
     End If
     XX5 = !HsInge_Otrab
     If XX5 > 0.005 Then
        Text19 = FORMATNUM(XX5, "F6.1")
     End If
     XX6 = !HsAvor_Otrab
     If XX6 > 0.005 Then
        Text4 = FORMATNUM(XX6, "F6.1")
     End If
     XX7 = !HsCCal_Otrab
     If XX7 > 0.005 Then
        Text14 = FORMATNUM(XX7, "F6.1")
     End If
     Text35 = !Observa_Otrab
     '
   End With
   '
   Call CARGOPLACO
   Call SUMAHS
   MODIFIC% = 0
   YACAR% = 1
   '
End Sub
'
Sub GENEINDIOTRAB() 'GENERA EL INDICE DE ORDENES DE TRABAJO
   'If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABREORTRA
      Set ORDETRATECO = OTrabajo.OpenRecordset("SELECT NRO_oTRAB, Nclie_Otrab, REFE_oTRAB,FEMI_oTRAB FROM ORDETRAB  ORDER BY REFE_OTRAB ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      ORDETRATECO.MoveLast
      If Err Then Exit Sub
      '
      CAREC% = ORDETRATECO.RecordCount
      If CAREC% < 1 Then Exit Sub
      '
      FIN& = CAREC%
      REBLB$ = REGBLAN$("INDIOTRA")
      JJ& = 0
      With ORDETRATECO
        .MoveFirst
        '
10      Call TRACE(20, JJ&, FIN&)
        JJ& = JJ& + 1
        NUCLIOTR% = !Nclie_Otrab
        EL2$ = REBLB$
        Call REPLA(EL2$, !Nro_Otrab, 1, 16)
        Call REPLA(EL2$, TRIM$(ABRECLI$(NUCLIOTR%)) & " - " & !Refe_Otrab, 17, 48)
        Call GRAREG("INDIOTRA", EL2$, JJ&)
        '
        If JJ& < CAREC% Then
          ORDETRATECO.MoveNext
          GoTo 10
        End If
      End With
      '
      ORDETRATECO.Close
      Call FILESORT("INDIOTRA", "", 1, 1, "ASC")
      Call FRESHECHO("INDIOTRA")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
   'End If
End Sub
'
Sub PRIORVEN()
    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
    'Call SETULTREG("!OBSERVOF", 0)
    '
    ' IMPRESION DE LA Especificación de Vta.
    '
    FORIPRE$ = CONTROL$("", "ESPVENTA")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Text1)
      TIPODOC$ = "Especificación de Venta"
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "IMP-TOTA" 'IMPORTE TOTAL
      CODPARAM$(2) = "VENDEDOR" 'REPRESENTANTE
      CODPARAM$(3) = "FECH-LAN" 'FECHA DE LA OFERTA
      CODPARAM$(4) = "FECH-REC" 'FECHA DEL PEDIDO
      CODPARAM$(5) = "MONEDEMI" 'MONEDA DE EMISION
      CODPARAM$(6) = "DOMI-TRA" 'LUGAR DE ENTREGA
      CODPARAM$(7) = "NRO-OCOM" 'NUMERO DE PEDIDO
      CODPARAM$(8) = "NRO-AENT" 'NUMERO DE OFERTA
      CODPARAM$(9) = "FECH-VEN" 'FECHA DE ENTREGA
      CODPARAM$(10) = "TIE-PREP" 'HS. INGENIERIA
      CODPARAM$(11) = "TIE-PROD" 'HS AVOR
      CODPARAM$(12) = "TIE-TOTA" 'HS. CCAL
      CODPARAM$(13) = "TOT-PREP" 'HS. CALDERERIA
      CODPARAM$(14) = "TOT-PROD" 'HS. MECANIZADO
      CODPARAM$(15) = "EQUIPO1" 'HS. MONTAJE INT.
      CODPARAM$(16) = "EQUIPO2" 'HS. MONTAJE EXT.
      CODPARAM$(17) = "REF-ASIE" 'USUARIO APROBACION
      CODPARAM$(18) = "TOT-TIEM" 'HS. TOTAL
      CODPARAM$(19) = "CANTIDAD" 'CANTIDAD
      CODPARAM$(20) = "PRE-NETO" 'PRECIO NETO
      CODPARAM$(21) = "NRO-PASE" 'COMISION
      
      CAPARDOC% = 21
      
      DOCPARAM(1) = Val(Text38)
      DOCPARAM(2) = DETEXT(5)
      DOCPARAM(3) = Text12
      DOCPARAM(4) = Text24
      DOCPARAM(5) = Label2
      DOCPARAM(6) = Text2
      DOCPARAM(7) = Text25
      DOCPARAM(8) = Text13
      DOCPARAM(9) = Text10
      DOCPARAM(10) = Text19
      DOCPARAM(11) = Text4
      DOCPARAM(12) = Text14
      DOCPARAM(13) = Text30
      DOCPARAM(14) = Text29
      DOCPARAM(15) = Text28
      DOCPARAM(16) = Text32
      DOCPARAM(17) = " "
      DOCPARAM(18) = Label44.Caption
      DOCPARAM(19) = Text36
      DOCPARAM(20) = Text38
      DOCPARAM(21) = Text21
      '
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 1
      REPLAORI% = 0
      RECOPROP% = 0
      REVICODX% = 0
      '
      CODTABU1$(1) = "DES-MAT"
      CACOLTAB1% = 1
      '
ANTEX% = 52
      Call FRATEXTO(Text17, ANTEX%)
      For UI% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = RETEX$(UI%)
      Next UI%
      '
      CODTABU2$(1) = "TIPO-VAL"
      CODTABU2$(2) = "BANC-CHE"
      CODTABU2$(3) = "IMPO-CHE"
      CODTABU2$(4) = "FECH-CHE"
      CACOLTAB2% = 4
      '
      CAITAB2% = 0
      For U% = 1 To List1.ListCount
        Y$ = List1.List(U% - 1)
        CAITAB2% = CAITAB2% + 1
        TETABU2$(1, CAITAB2%) = Mid$(Y$, 1, 3)
        TETABU2$(2, CAITAB2%) = Mid$(Y$, 4, 15)
        TETABU2$(3, CAITAB2%) = Mid$(Y$, 21, 10)
        TETABU2$(4, CAITAB2%) = Mid$(Y$, 33, 8)
      Next U%
      '
      Call FRATEXTO(Text35, 50)
      For KKU& = 1 To CARETEX%
       Call REGAPP("!OBSERVOF", RETEX$(KKU&))
      Next KKU&
      Call CIERRARCH("!OBSERVOF")

      Call PRINTDOC("IDORTRAB")                   ' Orden de TRABAJO
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    Call CIERRARCH("*.*")
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '

'    CPAG% = Val(Text21(0))
'    FORPAG$ = TRIM$(Eco21(0))
'    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
'    TIEM% = Val(Text21(1))
'    EMBALA$ = TRIM$(Eco21(1))
'    '
'    Call SETULTREG("!OBSERVOF", 0)
'    Call FRATEXTO(Text13, 52)
'    For KKU& = 1 To CARETEX%
'      Call REGAPP("!OBSERVOF", RETEX$(KKU&))
'    Next KKU&
'    Call CIERRARCH("!OBSERVOF")
'    '
'    PRIPLANOI% = 0
'    PRIESPEC% = 0
'    PRIHRINS% = 0
'    '
'    ' IMPRESION DE LA O/COMPRA
'    '
'    FORIPRE$ = CONTROL$("", "FORMOCOM")
'    '
'    If TRIM$(FORIPRE$) <> "" Then
'      '
'      Screen.MousePointer = 11
'      FECHEMII% = FECHANUM(Text9)
'      FECHDOC$ = FETEXTO$(FECHEMII%)
'      NUMEDOC$ = TRIM$(Text11)
'      While Len(NUMEDOC$) < 6
'        NUMEDOC$ = "0" + NUMEDOC$
'      Wend
'      TIPODOC$ = "Nota de Pedido"
'      '
'      CODPARAM$(1) = "COND-PAG"
'      CODPARAM$(2) = "NUBULTOS"
'      CODPARAM$(3) = "IMP-NETO"
'      CODPARAM$(4) = "POR-IVA"
'      CODPARAM$(5) = "IMP-IVA"
'      CODPARAM$(6) = "IMP-TOTA"
'      CODPARAM$(7) = "VENDEDOR"
'      CODPARAM$(8) = "LIST-PRE"
'      CODPARAM$(9) = "ZONA-VEN"
'      CODPARAM$(10) = "FECH-LAN"
'      CODPARAM$(11) = "IMPNEDES"
'      CODPARAM$(12) = "TRANSPOR"
'      CODPARAM$(13) = "POR-DESC"
'      CODPARAM$(14) = "IMP-DESC"
'      CODPARAM$(15) = "FECH-REC"
'      CODPARAM$(16) = "TIPO-VAL"
'      CODPARAM$(17) = "IMPO-CHE"
'      CODPARAM$(18) = "REF-CONJ"
'      CODPARAM$(19) = "DOMI-TRA"
'      CODPARAM$(20) = "PERC-IVA"
'      CAPARDOC% = 20
'      '
'      DOCPARAM$(1) = FORPAG$
'      DOCPARAM$(2) = EMBALA$
'      DOCPARAM$(3) = Total(0)
'      DOCPARAM$(4) = Text12
'      DOCPARAM$(5) = Total(2)
'      DOCPARAM$(6) = Total(3)
'      DOCPARAM$(7) = ""                           ' Mid$(X1$, 147, 20)
'      DOCPARAM$(8) = ""                           ' Mid$(X1$, 67, 40)
'      DOCPARAM$(9) = ""                           ' Mid$(X1$, 167, 40)
'      DOCPARAM$(10) = ""                          ' Mid$(X1$, 107, 40)
'      DOCPARAM$(11) = Str$(Val(Total(0)) - Val(Total(1)))
'      DOCPARAM$(12) = TRIM$(Text6)                ' ATENCION SR.
'      DOCPARAM$(13) = TRIM$(PORDESCU)
'      DOCPARAM$(14) = Total(1)
'      HOII% = HOY(HOS$)
'      FECHENT% = FECHANUM(Text14)
'      DOCPARAM$(15) = Text14
'      If FECHENT% = HOII% Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
'        ElseIf FECHENT% < HOII% + 4 Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
'      End If
'      DOCPARAM$(16) = TRIM$(Label11)
'      DOCPARAM$(17) = ""
'      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
'      DOCPARAM$(18) = Eco21(1)
'      DOCPARAM$(19) = Text10
'      DOCPARAM$(20) = Total(4)
'      '
'      CODTABU1$(1) = "COD-MAT"
'      CODTABU1$(2) = "DES-MAT"
'      CODTABU1$(3) = "UNIMED"
'      CODTABU1$(4) = "CANTIDAD"
'      CODTABU1$(5) = "FECH-VEN"
'      CODTABU1$(6) = "REF-MAT1"
'      CODTABU1$(7) = "PRE-LIST"
'      CODTABU1$(8) = "PORDESCU"
'      CODTABU1$(9) = "PRE-NETO"
'      CODTABU1$(10) = "NET-ITEM"
'      CODTABU1$(11) = "CAKILOS"
'      CODTABU1$(12) = "F-PLANO"
'      CODTABU1$(13) = "ORD-ITEM"
'      CACOLTAB1% = 13
'      '
'      DESTIDOC% = (-1) * Val(Text1)
'      CAITAB1% = 0
'      CAITCRU% = 0
'      INCLUSPE% = 1
'      REPLAORI% = 0
'      RECOPROP% = 0
'      REVICODX% = 0
'      '
'      NC1% = DESTIDOC%
'         CANTU = Val(Text20)                      ' CVS(Mid$(X$, 55, 4))
'         PREUNID# = Val(Text19)
'         PORDE = 0                                ' CVS(Mid$(X$, 75, 4))
'         FEX = FECHANUM(Text14)                   ' CVI(Mid$(X$, 95, 2))
'         CONSIMAT% = 0                            ' If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
'         '
'         CAITAB1% = CAITAB1% + 1
'         CODDX$ = TRIM$(Text7)
'         TETABU1$(1, CAITAB1%) = CODDX$
'           DECRIT$ = Text18                       ' TRIM$(DESCRIT0$(CI0%))
'         TETABU1$(2, CAITAB1%) = DECRIT$
'         TETABU1$(3, CAITAB1%) = Text22           ' Mid$(X$, 59, 4)
'         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
'         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
'         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  ' consigna m.prima
'         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
'         TETABU1$(8, CAITAB1%) = Str$(PORDE)
'         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
'         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
'           KILOX$ = ""                            ' FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
'         TETABU1$(11, CAITAB1%) = KILOX$
'         TETABU1$(12, CAITAB1%) = ""              ' FEPLANO$(CI0%)
'         TETABU1$(13, CAITAB1%) = "1"             ' TRIM$(Str$(J&))
'         '
'         If INCLUSPE% > 0 Then
'           TXTX$ = Text28
'           Call FRATEXTO(TXTX$, 63)
'           INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
'           '
'           COLUTA% = 0
'           URN& = CARETEX%
'           '
'           For U& = 1 To CARETEX%
'             TTXY$ = RTrim$(RETEX$(U&))
'             TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
'             If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
'                  CAITAB1% = CAITAB1% + 1
'                  TETABU1$(2, CAITAB1%) = TTYY$
'                Else
'                  COLUTA% = COLUTA% + 1
'                  If COLUTA% > 2 Then COLUTA% = 1
'                  If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
'                  TETABU1$(COLUTA%, CAITAB1%) = TTYY$
'             End If
'           Next U&
'           '
'         End If
'         '
'      Call PRINTDOC("FORMOCOM")                   ' Orden de Compra
'      '
'      For KK1% = 1 To CACOLTAB1%
'        For KK2% = 1 To CAITAB1%
'          TETABU1$(KK1%, KK2%) = ""
'        Next KK2%
'      Next KK1%
'      '
'    End If
'199 Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    Screen.MousePointer = 1
'    '
'    If EMICONSI% = 1 Then
'      If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
'        NPRO% = Val(Text1)
'        Call ENTRECONSIACC(NPRO%)
'      End If
'    End If
'    Call CIERRARCH("*.*")
'    '
'    Exit Sub
'    '
'999 Call CIERRARCH("*.*")
'    Screen.MousePointer = 1
'    '

End Sub

Sub CARGOPLACO() 'CARGO PLAN DE COBRA. DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM PLACOB "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With OTTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do Until .EOF = True
         NOITOT1$ = CStr(!NuOrItem)
         DESCRIOT1$ = CStr(!Refe_Pcob)
         IMPOT1$ = FORMATNUM(!Impo_Pcob, "F10.2")
         FEOT1$ = FECHATEX(CVINT(!FeVen_Pcob))
         Y$ = Space$(64)
         Call REPLA(Y$, AJUSTD$(NOITOT1$, 3), 1, 3)
         Call REPLA(Y$, AJUSTI$(DESCRIOT1$, 15), 6, 15)
         Call REPLA(Y$, IMPOT1$, 21, 10)
         Call REPLA(Y$, FEOT1$, 33, 8)
         List1.AddItem Y$
        .MoveNext
      Loop
      
      '
   End With
End Sub

Sub CARGOLISGRU() 'CARGO LISTA DE GRUPOS DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM GRUPOTRA "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   REBLA$ = REGBLAN$("GRUORTRA")
   JJ& = 0
   List3.Clear
   '
   With OTTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do Until .EOF = True
         NROGRU$ = !Nro_Grup
         If IsNull(!PROG_GRUP) Then
            PROGRU$ = ""
         Else
            PROGRU$ = !PROG_GRUP
         End If
         If IsNull(!COMP_GRUP) Then
            COMGRU$ = ""
         Else
            COMGRU$ = !COMP_GRUP
         End If
         If IsNull(!DENO_GRUP) Then
            DENGRU$ = ""
         Else
            DENGRU$ = !DENO_GRUP
         End If
         If IsNull(!OBSERVA_GRUP) Then
            OBSGRU$ = ""
         Else
            OBSGRU$ = !OBSERVA_GRUP
         End If
         '
         Z$ = REBLA$
         Call REPLA(Z$, NROGRU$, 1, 8)
         Call REPLA(Z$, PROGRU$, 9, 12)
         Call REPLA(Z$, COMGRU$, 21, 12)
         Call REPLA(Z$, DENGRU$, 33, 64)
         Call REPLA(Z$, OBSGRU$, 97, 64)
         JJ& = JJ& + 1
         Call GRAREG("!GRUORTRA", Z$, JJ&)
         '
         Y$ = Space$(128)
         Call REPLA(Y$, NROGRU$, 1, 8)
         Call REPLA(Y$, PROGRU$, 10, 6)
         Call REPLA(Y$, COMGRU$, 17, 6)
         Call REPLA(Y$, DENGRU$, 24, 48)
         Call REPLA(Y$, OBSGRU$, 73, 24)
         List3.AddItem Y$
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Call SETULTREG("!GRUORTRA", JJ&)
   '
End Sub

