VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FORTRA05 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Trabajo - A-B-M-L Base de Datos de O.T."
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FORTRA05.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
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
      TabIndex        =   73
      Top             =   840
      Width           =   870
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
      TabIndex        =   65
      Top             =   210
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
      Picture         =   "FORTRA05.frx":0442
      TabIndex        =   1
      Top             =   210
      Width           =   175
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
      TabIndex        =   52
      Top             =   525
      Width           =   3300
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
      TabIndex        =   51
      Top             =   525
      Width           =   1125
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
      TabIndex        =   49
      Top             =   840
      Width           =   1650
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
      Left            =   10750
      Picture         =   "FORTRA05.frx":074C
      TabIndex        =   47
      Top             =   525
      Width           =   175
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
      TabIndex        =   45
      Top             =   210
      Width           =   1125
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
      TabIndex        =   43
      Top             =   840
      Width           =   2340
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
      Left            =   10200
      TabIndex        =   37
      Top             =   7800
      Visible         =   0   'False
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "FORTRA05.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Picture         =   "FORTRA05.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Picture         =   "FORTRA05.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   38
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
      TabIndex        =   27
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
         Picture         =   "FORTRA05.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   80
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
         Picture         =   "FORTRA05.frx":26DA
         Style           =   1  'Graphical
         TabIndex        =   63
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
         Picture         =   "FORTRA05.frx":29E4
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
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
         Picture         =   "FORTRA05.frx":2B2E
         Style           =   1  'Graphical
         TabIndex        =   61
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
         TabIndex        =   35
         Top             =   7245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
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
         Picture         =   "FORTRA05.frx":2E38
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Imprime Ficha del Cliente"
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
         Picture         =   "FORTRA05.frx":34A2
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "FORTRA05.frx":37AC
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "FORTRA05.frx":3BEE
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "FORTRA05.frx":3EF8
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "FORTRA05.frx":4042
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "FORTRA05.frx":434C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   2310
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FORTRA05.frx":4496
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
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
      TabIndex        =   2
      Top             =   525
      Width           =   2925
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
      TabIndex        =   0
      Text            =   " "
      Top             =   210
      Width           =   1725
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   150
      TabIndex        =   8
      Top             =   1260
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   327681
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Especificacion de Venta"
      TabPicture(0)   =   "FORTRA05.frx":63B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Horas / Observaciones"
      TabPicture(1)   =   "FORTRA05.frx":63D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Lista de Grupos"
      TabPicture(2)   =   "FORTRA05.frx":63F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
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
         Height          =   3375
         Left            =   -74640
         TabIndex        =   104
         Top             =   2880
         Width           =   9975
         Begin VB.TextBox Text35 
            Height          =   2895
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   105
            Top             =   360
            Width           =   9735
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
         Height          =   2130
         Left            =   -69600
         TabIndex        =   88
         Top             =   525
         Width           =   4950
         Begin VB.TextBox Text34 
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
            Left            =   1130
            TabIndex        =   102
            Top             =   1680
            Width           =   1680
         End
         Begin VB.TextBox Text33 
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
            Left            =   3120
            TabIndex        =   101
            Top             =   1680
            Width           =   1680
         End
         Begin VB.TextBox Text32 
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
            Left            =   1130
            TabIndex        =   99
            Top             =   1220
            Width           =   1680
         End
         Begin VB.TextBox Text31 
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
            Left            =   3120
            TabIndex        =   98
            Top             =   1220
            Width           =   1680
         End
         Begin VB.TextBox Text30 
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
            Left            =   1130
            TabIndex        =   94
            Top             =   315
            Width           =   1680
         End
         Begin VB.TextBox Text29 
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
            Left            =   1130
            TabIndex        =   93
            Top             =   610
            Width           =   1680
         End
         Begin VB.TextBox Text28 
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
            Left            =   1130
            TabIndex        =   92
            Top             =   910
            Width           =   1680
         End
         Begin VB.TextBox Text27 
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
            Left            =   3120
            TabIndex        =   91
            Top             =   315
            Width           =   1680
         End
         Begin VB.TextBox Text26 
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
            Left            =   3120
            TabIndex        =   90
            Top             =   610
            Width           =   1680
         End
         Begin VB.TextBox Text23 
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
            Left            =   3120
            TabIndex        =   89
            Top             =   910
            Width           =   1680
         End
         Begin VB.Label Label29 
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
            Left            =   -840
            TabIndex        =   103
            Top             =   1725
            Width           =   1905
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Montaje Ext.:"
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
            Left            =   -750
            TabIndex        =   100
            Top             =   1280
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
            Left            =   -855
            TabIndex        =   97
            Top             =   360
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
            Left            =   -790
            TabIndex        =   96
            Top             =   645
            Width           =   1905
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Montaje Int.:"
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
            Left            =   -790
            TabIndex        =   95
            Top             =   960
            Width           =   1905
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Horas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   -74685
         TabIndex        =   66
         Top             =   525
         Width           =   4950
         Begin VB.TextBox Text22 
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
            Left            =   3120
            TabIndex        =   87
            Top             =   910
            Width           =   1680
         End
         Begin VB.TextBox Text20 
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
            Left            =   3120
            TabIndex        =   86
            Top             =   610
            Width           =   1680
         End
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
            Height          =   285
            Left            =   3120
            TabIndex        =   85
            Top             =   315
            Width           =   1680
         End
         Begin VB.TextBox Text14 
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
            Left            =   1095
            TabIndex        =   83
            Top             =   910
            Width           =   1680
         End
         Begin VB.TextBox Text4 
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
            Left            =   1095
            TabIndex        =   81
            Top             =   610
            Width           =   1680
         End
         Begin VB.TextBox Text19 
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
            Left            =   1080
            TabIndex        =   67
            Top             =   315
            Width           =   1680
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
            Left            =   -840
            TabIndex        =   84
            Top             =   960
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
            Left            =   -840
            TabIndex        =   82
            Top             =   645
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
            Left            =   -855
            TabIndex        =   68
            Top             =   360
            Width           =   1905
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   5880
         Left            =   -74600
         TabIndex        =   21
         Top             =   440
         Width           =   9990
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   5430
            Left            =   105
            ScaleHeight     =   5370
            ScaleWidth      =   9615
            TabIndex        =   22
            Top             =   525
            Width           =   9675
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   330
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   10515
               Begin VB.Line Line5 
                  X1              =   7320
                  X2              =   7320
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   "   Grupo           Prog.               Comp.                     Denominación                      Observaciones               "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   690
                  Left            =   120
                  TabIndex        =   24
                  Top             =   0
                  Width           =   10530
               End
               Begin VB.Line Line6 
                  X1              =   1320
                  X2              =   1320
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line7 
                  X1              =   2520
                  X2              =   2520
                  Y1              =   0
                  Y2              =   1050
               End
               Begin VB.Line Line13 
                  X1              =   4200
                  X2              =   4200
                  Y1              =   0
                  Y2              =   1050
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00C0F0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   6420
               Left            =   0
               TabIndex        =   25
               Top             =   -1050
               Width           =   10515
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
                  Height          =   5010
                  Left            =   -30
                  TabIndex        =   26
                  Top             =   1365
                  Width           =   9670
               End
               Begin VB.Line Line21 
                  X1              =   1155
                  X2              =   1155
                  Y1              =   1470
                  Y2              =   3255
               End
            End
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6120
         Left            =   200
         TabIndex        =   13
         Top             =   420
         Width           =   10320
         Begin VB.Frame Frame3 
            Caption         =   "Plan Cobranza"
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
            TabIndex        =   59
            Top             =   4515
            Width           =   5160
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   870
               Left            =   105
               TabIndex        =   79
               Top             =   600
               Width           =   4950
            End
            Begin VB.PictureBox Picture4 
               Height          =   225
               Left            =   105
               ScaleHeight     =   165
               ScaleWidth      =   4890
               TabIndex        =   77
               Top             =   315
               Width           =   4950
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "It.           Descripcion           Importe    Fecha"
                  BeginProperty Font 
                     Name            =   "Lucida Console"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   0
                  TabIndex        =   78
                  Top             =   0
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
            Left            =   5250
            TabIndex        =   57
            Top             =   4515
            Width           =   2640
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
               TabIndex        =   58
               Top             =   315
               Width           =   2430
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
            TabIndex        =   16
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
               TabIndex        =   7
               Text            =   " "
               Top             =   1155
               Width           =   975
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
               Left            =   2100
               TabIndex        =   6
               Top             =   420
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
               Left            =   1515
               TabIndex        =   5
               Text            =   " "
               Top             =   420
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
               TabIndex        =   19
               Top             =   735
               Width           =   2145
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
               TabIndex        =   18
               Top             =   1230
               Width           =   1710
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor:"
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
               Left            =   -735
               TabIndex        =   17
               Top             =   525
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
            TabIndex        =   14
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
               Picture         =   "FORTRA05.frx":640C
               TabIndex        =   72
               Top             =   315
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
               Picture         =   "FORTRA05.frx":6716
               TabIndex        =   71
               Top             =   315
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
               TabIndex        =   54
               Top             =   315
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
               TabIndex        =   53
               Text            =   " "
               Top             =   315
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
               TabIndex        =   4
               Text            =   " "
               Top             =   315
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
               TabIndex        =   3
               Top             =   315
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
               TabIndex        =   56
               Top             =   420
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
               TabIndex        =   55
               Top             =   420
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
               TabIndex        =   41
               Top             =   420
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
               TabIndex        =   15
               Top             =   420
               Width           =   1905
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -945
            TabIndex        =   60
            Top             =   840
            Width           =   11670
            Begin VB.TextBox Text38 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
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
               Left            =   9600
               TabIndex        =   108
               Top             =   600
               Width           =   1555
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
               TabIndex        =   107
               Text            =   "Text36"
               Top             =   0
               Width           =   825
            End
            Begin VB.TextBox Text36 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
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
               Left            =   1043
               TabIndex        =   106
               Top             =   600
               Width           =   765
            End
            Begin VB.TextBox Text6 
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
               Height          =   2595
               Left            =   9555
               TabIndex        =   76
               Top             =   525
               Width           =   1630
            End
            Begin VB.PictureBox Picture2 
               Height          =   330
               Left            =   1020
               ScaleHeight     =   270
               ScaleWidth      =   10095
               TabIndex        =   74
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
                  Caption         =   $"FORTRA05.frx":6A20
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
                  Left            =   0
                  TabIndex        =   75
                  Top             =   0
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
            Begin VB.TextBox Text15 
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
               Height          =   2955
               Left            =   1020
               TabIndex        =   70
               Top             =   525
               Width           =   870
            End
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
               Height          =   2955
               Left            =   1890
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   64
               Top             =   525
               Width           =   7680
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
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
               Left            =   9600
               TabIndex        =   109
               Top             =   3120
               Width           =   1575
            End
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   12
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   11
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   330
      Left            =   3360
      TabIndex        =   110
      Top             =   105
      Width           =   750
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
      TabIndex        =   69
      Top             =   315
      Width           =   1110
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
      TabIndex        =   50
      Top             =   945
      Width           =   585
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
      TabIndex        =   48
      Top             =   630
      Width           =   810
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
      TabIndex        =   46
      Top             =   315
      Width           =   690
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
      TabIndex        =   44
      Top             =   945
      Width           =   900
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
      TabIndex        =   42
      Top             =   630
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
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
      TabIndex        =   10
      Top             =   630
      Width           =   1350
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
      TabIndex        =   9
      Top             =   315
      Width           =   1140
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
      Begin VB.Menu DeHaCaCli 
         Caption         =   "Deshacer Cambios"
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
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConPagCli 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu LiPreCli 
         Caption         =   "Listas de Precios"
      End
      Begin VB.Menu VenCli 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu PrZoVeCli 
         Caption         =   "Provincias y Zonas de Venta"
      End
      Begin VB.Menu GruCli 
         Caption         =   "Grupo de Clientes"
      End
      Begin VB.Menu CuCoCli 
         Caption         =   "Cuentas Contables"
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FORTRA05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$


Dim UPDATOTR% 'MARCA SI CAMBIA ALGUN DATO

Private Sub AyuCli_Click()
Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
   Call Command24_Click
End Sub

Private Sub BusCli_Click()
Call Command20_Click
End Sub

Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub Command11_Click()
   '
   Call COMUNI("No Disponible en la presente versión")
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
Call COMUNI("No Disponible en la presente versión")
'   Static CTXX%
'   Command22.Enabled = False
'   '
'   If ULTREG&("DAPIBCLI") < 1 Then
'     Call GENDAPIBCLI
'   End If
'   '
'   NCLIE% = Val(TRIM$(Text1.TEXT))
'   If NCLIE% > 0 Then
'     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
'       '
'       If CTXX% < 1 Then
'         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
'         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
'         Call SETULTREG("TACATIBC", 2)
'         CTXX% = 1
'       End If
'       '
'       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
'       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
'       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
'       If CATIB1% < 0 Or CATIB1% > 1 Then
'         Call REPLA(RECORD$, Chr$(1), 3, 1)
'         Call REPLA(RECORD$, Space$(32), 9, 32)
'       End If
'       '
'10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
'       If OBX$ <> "" Then
'         RECORD$ = OBX$
'         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
'         If CATIB1% < 0 Or CATIB1% > 1 Then
'           Call MENSERR(24, "Mal Categoría I.Brutos")
'           GoTo 10
'         End If
'         If CATIB1% = 0 Then
'           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
'             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
'             GoTo 10
'           End If
'         End If
'         If CATIB1% = 1 Then
'           Call REPLA(RECORD$, Space$(32), 9, 32)
'         End If
'         PORCONVE = CVS(Mid$(RECORD$, 5, 4))
'         If PORCONVE < 0 Or PORCONVE > 100 Then
'           Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
'           GoTo 10
'         End If
'         '
'         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
'         Call REPLA(RECORD$, USERNAME$, 41, 20)
'         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
'         Call REPLA(RECORD$, MKI$(HORANUM%(Time$)), 63, 2)
'         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
'         Call CIERRARCH("DAPIBCLI")
'         '
'       End If
'     End If
'   End If
'   Command22.Enabled = True
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
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

   Call SELECHO("VENDEDOR")
   VDEVU$ = VALACT1$("VENDEDOR")
   If TRIM$(VDEVU$) <> "" Then
     COTEXT(5).TEXT = VDEVU$
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENCLI
   Call GRANOTACLI
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
   Call COMUNI("No Disponible en la presente versión")
   'AMACLIS04.Show 1
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPORTRA04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    Call GRABOT
    Call SUMAHS
End Sub
Sub SUMAHS()
    XX1 = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
    Text34 = FORMATNUM(XX1, "F6.2")
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENCLI
   Call GRANOTACLI
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATCLIE
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
   If TRIM$(NUOTR$) = "" Then
      Call COMUNI("Número de Orden de Trabajo no Válida o Pre-Existente")
      GoTo 9
   End If
'   'ADEMAS HAY QUE VALIDAR QUE YA NO EXISTA EL NUMERO
'   'la sig. validacion falta testear
'   Call ABREORTRA
'   ORDETRAB.FindFirst "Nro_Otrab = " & NuOTR$
'   If ORDETRAB.NoMatch = False Then
'      Call COMUNI("Numero de Orden de Trabajo no Válido o Pre-Existente")
'      GoTo 9
'   End If
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
   '//////FIN DE VALIDACIONES////////
   'HASTA ACA OK
   '/////ALTA DEL REGISTRO//////
   Call ABREORTRA
   With ORDETRAB
      .AddNew
        !Nro_Otrab = NUOTR$
        !Refe_Otrab = NUOTR$ & "-" & DESCRIOTR$
        !NCLIE_oTRAB = NUCLIOTR%
        !FEMI_oTRAB = CVDAT(FECOTR%)
        !Fentre_Otrab = CVDAT(FEENOTR%)
        !uSUCONF_OTRAB = USERNAME 'usuario de confección
        !UsuCont_Otrab = "" 'usuario que controlo
        !UsuApro_Otrab = "" 'usuario que aprobara
        !NroOcom_Otrab = "" 'numero de orden de compra del cliente
        !FeOcom_Otrab = Date 'fecha de pedido del cliente
        !NroOfer_Otrab = "" 'Numero de Oferta
        !FeOFer_Otrab = Date 'fecha de la oferta
        !Nuoritem_Otrab = 0
        !CanPed_Otrab = 0
        !Especi_oTRAB = ""
        !preven_Otrab = 0
        !mone_Otrab = 5
        !CotiMone_Otrab = 0
        !Lentrega_Otrab = " "
        !Ajusta_Otrab = 0
        !Vended_Otrab = 0
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
   '//////////FIN DE ALTA//////////////////
   'ACA VIENE LA RUTINA PARA LEVANTAR A LA FICHA LA ORDEN DE TRABAJO.
   '
   Call GENEINDIOTRAB
   Text1 = TRIM$(NUOTR$)
   '
'   NCLIE% = NROPRO&
'   VDEF$ = String$(128, 0)
'   Call REPLA(VDEF$, MKI$(NCLIE%), 13, 2)
'   Call REPLA(VDEF$, Space$(62), 15, 62)
'   Call REPLA(VDEF$, Chr$(1), 77, 1)
'   Call REPLA(VDEF$, Space$(15), 78, 15)
'   '
'10 Screen.MousePointer = 1
'   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
'   If OBX$ = "" Then GoTo 99
'   VDEF$ = OBX$
'   '
'   COCLINU$ = TRIM$(Left$(OBX$, 12))
'   If COCLINU$ = "" Then
'     Call MENSERR(24, "Codigo de Cuenta no Válido")
'     GoTo 10
'   End If
'   '
'   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
'   If NCLINUE% < 1 Or NCLINUE% > 9999 Then
'      Call MENSERR(24, "Número de Cuenta no Válido.")
'      GoTo 10
'   End If
'   '
'   Set Tranclie = Badaclie.OpenRecordset("SELECT Codi_Cli FROM Deudor12;", dbOpenSnapshot)
'   On Error Resume Next
'   Tranclie.FindFirst "Codi_Cli = '" & COCLINU$ & "'"
'   If Tranclie.NoMatch = False Then
'      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
'      GoTo 10
'   End If
'   '
'   Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12;", dbOpenSnapshot)
'   On Error Resume Next
'   Tranclie.FindFirst "Nume_Cli = " & NCLINUE
'   If Tranclie.NoMatch = False Then
'      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
'      GoTo 10
'   End If
'   '
'   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
'   If Len(DCLINUE$) < 1 Then
'      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
'      GoTo 10
'   End If
'   '
'   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
'   If PIVCLI% < 0 Or PIVCLI% > 6 Then
'      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
'      GoTo 10
'   End If
'   '
'   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
'   If VALICUIT%(CUTT$) <> 1 Then
'     If PIVCLI% <> 0 Then
'       If PIVCLI% <> 5 Then
'         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
'         GoTo 10
'       End If
'     End If
'   End If
'   '
'   If CUTT$ <> "" Then
'     Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Cuit_Cli = '" & CUTT$ & "';", dbOpenSnapshot)
'     CAREI% = 0
'     On Error Resume Next
'     Tranclie.MoveLast
'     CAREI% = Tranclie.RecordCount
'     '
'     If CAREI% > 0 Then
'       If Tranclie!Nume_Cli > 0 Then
'         NCLIX% = Tranclie!Nume_Cli
'         If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
'           GoTo 10
'         End If
'       End If
'     End If
'   End If
'   '
'   Screen.MousePointer = 11
'20 Deudor12.Close
'   On Error Resume Next
'   Call ABRECLIEN
'   Deudor12.MoveFirst
'   If Err < 1 Then
'     On Error GoTo 0
'     ABUS$ = "Nume_Cli = " & NCLINUE%
'     Deudor12.FindFirst ABUS$
'     If Deudor12.NoMatch = False Then
'       NCLINUE% = NCLINUE% + 1
'       GoTo 20
'     End If
'   End If
'   '
'   With Deudor12
'      .AddNew
'      !Nume_Cli = TRIM$(Str$(NCLINUE%))
'      !Raso_Cli = DCLINUE$
'      !Codi_Cli = COCLINU$
'      !Cuit_Cli = CUTT$
'      !Piva_Cli = PIVCLI%
'      !Stat_Cli = 1
'      '
'      On Error Resume Next
'      .Update
'      If Err Then
'         Call MENSERR(24, "Alta no Procesada Correctamente")
'      End If
'      On Error GoTo 0
'      .Bookmark = .LastModified
'      .Close
'      UPDATOTR% = 1
'      Call GENECOCLI
'   End With
'   '
'   AGRECLI% = 1
'   Text1.TEXT = TRIM$(Str$(NCLINUE%))
'   SSTab1.Tab = 0
'   Call CARDATCLIE
'   Call CARTAVALI
'   '
'   For Index = 0 To 7
'     If Index <> 1 Then
'       VAPRE$ = TRIM$(CONTROL$("AMACLI", ARCHE$(Index)))
'       COTEXT(Index) = "": DETEXT(Index) = ""
'       If VAPRE$ <> "" Then
'         TV11% = TIPFIELD(ARCHE$(Index), 1)
'         LF11% = LENFIELD(ARCHE$(Index), 1)
'         ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
'         ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
'         If ECO$ <> "" Then
'             COTEXT(Index).TEXT = VAPRE$
'             DETEXT(Index).Caption = ECO$
'           Else
'             COTEXT(Index).TEXT = ""
'             DETEXT(Index).Caption = ""
'         End If
'       End If
'     End If
'   Next Index
'   '
'   Text3 = DCLINUE$
'   COTEXT(1) = TRIM$(Str$(PIVCLI%))
'   Text33 = TRIM$(CUTT$)
'   '
  Screen.MousePointer = 1
'   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   SSTab1.Tab = 0
   Text20.TEXT = " Suspendido"
   Text20.Refresh
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   If DERACCE%("BAJACLI", "Baja de Cuenta de Cliente") < 2 Then
     GoTo 99
   End If
   '
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE& = Val(Text1.TEXT)
       Call ABRECLIEN
       With Deudor12
       On Error Resume Next
       .MoveFirst
       .FindFirst "Nume_Cli = " & NCLIE&
       If .NoMatch = False Then
         '
         On Error GoTo 0
         NC% = NCLIE
         '
         If Abs(SALDCLI#(NC%)) >= 0.005 Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Saldo.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         
         TTXX$ = "La Baja de un Cliente\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Cobranza o Pedidos."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
           '
           .Edit
           !Stat_Cli = (-1)
           .Update
           '
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTACLI
           '
           UPDATOTR% = 1
           Call GENECOCLI
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
           '
         End If
         '
       End If
       End With
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
   Call PRIORVEN
  ' PRINTFORM
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
   '
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

Private Sub Label5_Click()
    Call List1_DblClick
End Sub

Private Sub LiGeClaCli_Click()
 Call AMACLIS04.Command8_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub List1_DblClick()
    '
    If TRIM$(Text1.TEXT) <> "" Then
        Call GENEPLACO 'GENERO EL DAT
        PRF$ = "CAPLACO"
        ATRI$ = "/NC"
        ATRI$ = ATRI$ + "/TITUPAN=Plan de Cobranza"
        VTB% = VENTABU%(PRF$ + ATRI$)
        If VTB% >= 0 Then
            Call GRACAPLACO ' GUARDO LOS REGISTROS EN TABLA
        End If
    End If
    '
End Sub

Sub GRACAPLACO()
    '
    'ACA VA RUTINA DE GRABACION
    Call ABREPLACO
    J& = ULTREG("CAPLACO")
    For KKIL& = 1 To J&
       XX$ = REGLEIDO$("CAPLACO", KKIL&)
       NUOTRAB$ = TRIM$(Text1.TEXT)
       ITE% = CVI(Mid$(XX$, 17, 2))
       DESCRI$ = Mid$(XX$, 19, 64)
       IMPOT# = CVD(Mid$(XX$, 83, 8))
       FECOTR% = CVI(Mid$(XX$, 91, 2))
       NUFAC = Mid$(XX$, 93, 16)
       With PLACOB
         .AddNew
           !Nro_Otrab = NUOTRAB$
           !NuOrItem = ITE%
           !Refe_Pcob = DESCRI$
           !Impo_Pcob = IMPOT#
           !FeVen_Pcob = CVDAT(FECOTR%)
         .Update
       End With
    Next KKIL&
    '
End Sub
Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREPLACO
    SQLX$ = "SELECT * FROM PLACOB "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTRABAJO.OpenRecordset(SQLX$, 4)
    
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

Private Sub Text11_GotFocus()
    Text25.SetFocus
End Sub


Sub GRABOT() 'GRABAR ORDEN DE TRABAJO
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
   If TRIM$(Text3) = "" Then
      Call COMUNI("Descripción de Orden de Trabajo no válido.")
      Exit Sub
   End If
   '
   If CDbl(CVINT(CDate(Text10))) < CDbl(CVINT(CDate(Text10))) Then
       Call COMUNI("Fecha de Entrega no válida.")
       Exit Sub
   End If
   If TRIM$(Text10) = "" Then
       Call COMUNI("Fecha de Entrega no válida.")
       Exit Sub
   End If
   '
   If TRIM$(Text17) = "" Then
       Call COMUNI("Especificación de Venta Incompleta.")
       Exit Sub
   End If
   '
   If Val(TRIM$(Text36)) <= 0 Then
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
   Set OTTMP = OTRABAJO.OpenRecordset(SQLX$, 2)
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
        !FeOFer_Otrab = CVDAT(CVINT(Text12))
        !CanPed_Otrab = Val(Text36)
        !preven_Otrab = Val(Text38)
        !Lentrega_Otrab = Text2
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
            !Vended_Otrab = Val(COTEXT(5))
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
        '
     .Update
    End With
   '
End Sub

Private Sub Text14_LOSTFOCUS()
    Text14 = FORMATNUM(Val(Text14), "F6.1")
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

Private Sub Text38_lostfocus()
    '
    Text38 = FORMATNUM$(Val(Text38), "F12.2")
    '
End Sub

Private Sub Text4_LOSTFOCUS()
    Text4 = FORMATNUM(Val(Text4), "F6.1")
End Sub

Private Sub Text5_GotFocus()
    Text25.SetFocus
End Sub

Private Sub Text6_Change()
   MODIFIC%(0) = 1
   If NUMECLI%(Text6) > 0 Then
     Text1 = TRIM$(Str$(NUMECLI%(Text6)))
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command10_Click
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
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command10_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Screen.MousePointer = 1
   RSLX$ = REGSEL$("INDICLIE")
   If Len(RSLX$) > 4 Then
     NCLIE% = CVI(Mid$(RSLX$, 63, 2))
     If NCLIE% > 0 Then
       NC% = NCLIE%
       Text1.TEXT = TRIM$(Str$(NC%))
     End If
   End If
   '
End Sub

Sub CARDATCLIE()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE < 1 Or Abs(STCL%) < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text23.TEXT = NPROCLI$(NCLIE)
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENCLI
      Case 1
        Call CARNOTACLI
      Case 2
     '   Call CARESCUENTA
    End Select
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENCLI()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = Val(TRIM$(Text1.TEXT))
    '
    Text6 = CODICLI$(NC%)
    Text25.TEXT = DOMICLI$(NC%)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NC%)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NC%)
    Text27.Refresh
    If LIMICRE(NC%) < 0 Then
        Text20.TEXT = " Suspendido"
      Else
        Text20.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 12, 2, 2)
    End If
    Text20.Refresh
    Text23.TEXT = NPROCLI$(NC%)
    Text23.Refresh
    Text29.TEXT = PAISCLI$(NC%)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NC%)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NC%)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NC%)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NC%)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NC%)
    Text36.Refresh
    Text33.TEXT = CUITCLI$(NC%)
    Text33.Refresh
    Text34.TEXT = NIBRCLI$(NC%)
    Text34.Refresh
    Text21.TEXT = TRIM$(DESCUCLI$(NC%))
    Text21.Refresh
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    COTEXT(2).Refresh
    COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    COTEXT(3).Refresh
    COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    COTEXT(4).Refresh
    COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NC%)))
    COTEXT(5).Refresh
    COTEXT(6).TEXT = TRIM$(CUEMADRE$(NC%))
    COTEXT(6).Refresh
    COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NC%)))
    COTEXT(7).Refresh
    '
    'grupo divisional
    GRUDI% = GRUDIVCLI%(NC%)
    Text2 = ""
    If GRUDI% <> NC% Then Text2 = TRIM$(Str$(GRUDI%))
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENCLI()
    '
    NCLIE& = Val(TRIM$(Text1.TEXT))
    If NCLIE& < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Número de Cliente no Válido.")
      End If
      Exit Sub
    End If
    '
    If Abs(STATCLI%(NCLIE)) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Numero de Cliente Inexistente")
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If STATCLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        If COMALTER%("El Cliente Elegido Figura\Marcado Como 'Dado de Baja'.\  \Desea Re-Activar Esta Cuenta ?\\No, Dejar de Baja\Si, Re-Activar") = 2 Then
          Call ABRECLIEN
          With Deudor12
            On Error Resume Next
            .FindFirst "Nume_Cli = " & Abs(NCLIE&)
            If .NoMatch = False Then
             .Edit
             !Stat_Cli = 1
             .Update
            End If
          End With
          UPDATOTR% = 1
          Call GENECOCLI
          Call CARGENCLI
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Exit Sub
      End If
    End If
    '
    If GRABATODO% = 1 Then
      If SSTab1.Tab <> 0 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATCLI", "Modificar Datos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      Command9.SetFocus
      If COMALTER%("Se han Modificado Datos del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = RASOCLI$(NCLIE)
        Text24.TEXT = TEL2CLI$(NCLIE)
        MODIFIC(0) = 0
        Call CARGENCLI
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(Text33.TEXT)
    If CUIT$ <> "" Then
      If VALICUIT%(CUIT$) <> 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
        SSTab1.Tab = 0
        Text33.SetFocus
        Exit Sub
      End If
    End If
    '
    If NCLIE = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Imposible Registrar !!!\Falta Código de Cuenta")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NPRII% = NUMECLI%(Text6)
    If NPRII% > 0 Then
      If NPRII% <> Abs(NCLIE) Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Codigo de Cuenta Repetido o no Válido.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social del Cliente.")
        Exit Sub
    End If
    '
    '
    Call ABRECLIEN
    With Deudor12
      On Error Resume Next
      .MoveFirst
      .FindFirst "Nume_Cli = " & NCLIE&
      If .NoMatch = False Then GoTo 10
      On Error GoTo 0
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End With
    '
10  On Error GoTo 0
    With Deudor12
      On Error Resume Next
      ABC1$ = !Codi_Cli
      If Err Then
        ABC1$ = ""
      End If
      On Error GoTo 0
      .Edit
        If TRIM$(Text3) <> TRIM$(!Raso_Cli) Or UCase$(Text6) <> ABC1$ Then UPDATOTR% = 1
      !Codi_Cli = UCase$(Text6)
      !Raso_Cli = Text3
      !Stat_Cli = 1
      !Domi_Cli = Text25
      !Cpos_Cli = Text26
      !Loca_Cli = Text27
      !Tele_Cli = Text30
      !Fant_Cli = Text24
      !Cuit_Cli = TRIM$(CUIT$)
'      !Nibr_Cli = TRIM$(Text34)
      !Caib_Cli = TRIM$(COTEXT(2))
      !Prov_Cli = TRIM$(COTEXT(0))
      '
      MODLIMCRE% = 0
      On Error Resume Next
      LMCRA = CDbl(!Lcre_Cli)
      If Err Then
        LMCRA = 0: !Lcre_Cli = 0
      End If
      On Error GoTo 0
      '
      If Abs(CDbl(!Lcre_Cli) - Val(Text20.TEXT)) > 0.05 Or InStr(UCase$(Text20.TEXT), "SUSPEN") > 0 Then
        If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
            !Lcre_Cli = Val(Text20)
            If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
              !Lcre_Cli = (-1)
            End If
          Else
           Text20.TEXT = CSTRING$(MKS$(!Lmcr_Cli), 3, 12, 2, 2)
           If LMCR < 0 Then Text20.TEXT = " Suspendido"
           Text20.Refresh
           MODLIMCRE% = 1
        End If
      End If
      '
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      !Piva_Cli = PIVA%
      '
      CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
      !Caib_Cli = CATIB%
      '
      LIPRE% = VALIMI(COTEXT(3).TEXT, 0, 255)
      If LIPRE% > 0 Then
        If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
          Call MENSERR(24, "Lista de Precios no Válida")
          LIPRE% = 0
        End If
      End If
      !Lpre_Cli = LIPRE%
      '
      COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
      If COPAGI% > 0 Then
        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
        End If
      End If
      !Cpag_Cli = COPAGI%
      '
      VENDE% = VALIMI(COTEXT(5).TEXT, 0, 255)
      If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
          Call MENSERR(24, "Vendedor no Válido")
          VENDE% = 0
        End If
      End If
      !Vend_Cli = VENDE%
      '
      CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
      If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
        CUECO% = 0
      End If
      !GrCo_Cli = CUECO%
      !Gru_Div = Val(Text2)
      '
      CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(COTEXT(6).TEXT, 12))
      !Cuma_Cli = CCX%
      '
      DCC$ = TRIM$(Text21.TEXT)
      !Dsta_Cli = TRIM$(Text21)
      '
      !Pais_Cli = Text29
      !Faxi_Cli = Text31
      !Mail_Cli = Text32
      !Ctac_Cli = Text19
      !Prvd_Cli = Text23
      !Hora_Cli = Text36
      '
      .Update
    End With
    '
    Deudor12.Close
    Call GENECOCLI
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTACLI()
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Text28 = ""
   '
   If NCLIE% > 0 Then
     If REGICLI(NCLIE%) > 0 Then
       Call ABRECLIEN
       Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NCLIE%, dbOpenSnapshot)
       '
       On Error Resume Next '
       ClienTemp.MoveFirst
       If Err < 1 Then
         Text28 = ClienTemp!Nota_Cli
       End If
     End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTACLI()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   '
   If REGICLI(NCLIE%) < 1 Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("ANOTACLI", "Anotaciones de de Clientes") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     Command9.SetFocus
     If COMALTER%("Se han Modificado las Notas del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       Call CARNOTACLI
       Exit Sub
     End If
     Command9.SetFocus
   End If
   '
   Call ABRECLIEN
   Screen.MousePointer = 11
   TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
   '
   With Deudor12
     On Error Resume Next
     .MoveFirst
     .FindFirst "Nume_Cli = " & Abs(NCLIE%)
     If .NoMatch = False Then GoTo 10
     On Error GoTo 0
     Call MENSERR(24, "Imposible Registrar")
     Screen.MousePointer = 1
     Exit Sub
   End With
   '
10 On Error GoTo 0
   With Deudor12
     .Edit
     !Nota_Cli = TEXTO$
     .Update
   End With
   Deudor12.Close
   MODIFIC%(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
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
   Text34.TEXT = ""
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
Call Command1_Click
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
'   VUELBUS% = 1
'   '
'3  RELIS$ = List3.TEXT
'   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
'   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
'   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
'   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
'     NUDOCU$ = TRIM$(NUDOCU$)
'     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
'   End If
'   RACLI$ = TRIM$(Left$(RASOCLI$(Val(Text1.TEXT)), 8))
'   Screen.MousePointer = 11
'   For UI& = ULTREG&("PDOCLST") To 1 Step -1
'     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'     If InStr(XX$, TIDOCUM$) > 4 Then
'       If InStr(XX$, NUDOCU$) > 4 Then
'         If InStr(XX$, RACLI$) > 16 Or VUELBUS% = 2 Then
'           DOCUNU& = CVS(Left$(XX$, 4))
'           GoTo 5
'         End If
'       End If
'     End If
'   Next UI&
'   '
'   If VUELBUS% < 2 Then
'     VUELBUS% = 2
'     GoTo 3
'   End If
'   '
'   Screen.MousePointer = 1
'   Call COMUNI("Documento no Encontrado")
'   Exit Sub
'   '
'5  For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   '
'   REDOCU$ = REGACT$("PDOCLST")
'   URE& = ULTREG&("PDOCSPL")
'   LI& = 0: LS& = URE&
'6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'   GoTo 8
'   '
'7  Screen.MousePointer = 1
'   Call MENSERR(24, "Documento no Registrado")
'   Exit Sub
'   '
'8  TPSP$ = ""
'   Screen.MousePointer = 11
'   While L& > 1
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then GoTo 9
'     L& = L& - 1
'   Wend
'   '
'9  CAPAGINAS% = 0
'   For KKL& = L& To URE&
'     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'     End If
'   Next KKL&
'   '
'95 SPOOLSTRING$ = TPSP$
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   ALTUPAGINA = 7170
'   TOPEPAGINA = 0
'   PAGINACTIVA% = 0
'   CAPAGINAS% = 0
'   '
'   FIN& = Len(TPSP$)
'10 PPXX% = InStr(TPSP$, Chr$(255))
'   If PPXX% > 0 Then
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
'   End If
'   '
'   Screen.MousePointer = 1
'   PAGINACTIVA% = 0
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Visible = False
'     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'   Next KKI%
'   '
'   EDITFORM.Picture1(PAGINACTIVA%).Refresh
'   EDITFORM.VScroll1.Min = 0
'   EDITFORM.VScroll1.Value = 0
'   EDITFORM.VScroll1.Max = 1
'   If ALTUPAGINA > EDITFORM.Frame1.Height Then
'      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'   End If
'   EDITFORM.Command3.Enabled = False
'   EDITFORM.Command5.Enabled = False
'   EDITFORM.Show 1
'   '
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
         Call CARGOTR(NUOTR$)
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRAGENCLI
   'Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub
Sub text19_lostfocus()
   Text19 = FORMATNUM(Val(Text19), "F6.1")
End Sub
Private Sub Text20_Change()
   MODIFIC%(0) = 1
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
   MODIFIC%(0) = 1
End Sub



Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub
Sub TEXT28_LOSTFOCUS()
    Text28 = FORMATNUM(Val(Text28), "F6.1")
End Sub
Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub
Sub TEXT29_LOSTFOCUS()
    Text29 = FORMATNUM(Val(Text29), "F6.1")
End Sub
Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub
Sub TEXT30_LOSTFOCUS()
    Text30 = FORMATNUM(Val(Text30), "F6.1")
End Sub
Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub
Sub TEXT32_LOSTFOCUS()
    Text32 = FORMATNUM(Val(Text32), "F6.1")
End Sub
Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
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
Sub CARESCUENTA()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        List3.Clear
        List3.Refresh
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        LTEXT37.Caption = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 11, 2, 2)
        LMCR = LIMICRE(NC%)
        LTEXT36.Caption = CSTRING$(MKS$(LMCR), 4, 11, 2, 2)
        If LMCR < 0 Then LTEXT36.Caption = " Suspendido"
        Refresh
        LTEXT35.Caption = CSTRING$(MKD$(CHEPEACRE#(NC%)), 4, 11, 2, 2)
        DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
        LTEXT23.ForeColor = &H80000008
        If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
        LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
        Refresh
    End If
    '
100 MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    For U& = 1 To Len(RFC$) Step 4
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        FF% = CVI(Mid$(X$, 7, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$((-1) * REG&), 4, 8, 0, 2)
        LINDICE.AddItem Z$
    Next U&
    '
    SALDO# = 0: SARRAS# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& > MAXE& Then GoTo 990
            X$ = REGLEIDO$("FACTUR", REG&)
            TI% = Asc(Mid$(X$, 13, 1))
            TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                               If TI% = 1 Then TIPOS$ = "Venta Contado"
                               If TI% = 3 Then TIPOS$ = "Nota de Debito"
                               If TI% = 5 Then TIPOS$ = "Descuento Caja"
                               If TI% = 6 Then TIPOS$ = "Nota de Credito"
                               If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                               '
            NROL& = CVS(Mid$(X$, 3, 4))
            If NROL& < 1 Then
               If TI% = 0 Or TI% = 7 Then
                  Call REPLA(X$, MKS$(REG&), 3, 4)
                  Call GRAREG("FACTUR", X$, REG&)
               End If
            End If
            '
            SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
            IMPO# = CVD(Mid$(X$, 15, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
            IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
            IMPO# = SG% * IMPO#
            GoTo 890
            '
         ElseIf REG& < 0 Then
            REG1& = Abs(REG&)
            If REG1& > MAXDIAC& Then GoTo 990
            X$ = REGLEIDO$("COBRAN", REG1&)
            TI% = 4
            TIPOS$ = "Recibo"
            SG% = (-1)
            IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
            GoTo 890
        End If
        GoTo 990
        '
890     SALDO# = SALDO# + (Int(100 * IMPO# + 0.5)) / 100
        FEX = CVI(Mid$(X$, 7, 2))
        NROL& = CVS(Mid$(X$, 3, 4))
        Z$ = Space$(80)
        '
        Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
        Call REPLA(Z$, TIPOS$, 11, 16)
        Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
        IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
        SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
        If IMPO# > 0 Then
            Call REPLA(Z$, IMPTEX$, 35, 13)
          ElseIf IMPO# < 0 Then
            Call REPLA(Z$, IMPTEX$, 50, 13)
        End If
        Call REPLA(Z$, SALTEX$, 63, 14)
        List3.AddItem Z$
        '
990 Next U&
    '
991 SARRAS# = SALDO#
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        NCII% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) < 0.005 Then
          If Abs(SALDO#) > 0.0005 Then
            SALDO# = 0
            Call REPLA(X$, MKD$(0), 11, 8)
            Call GRAREG("SVD202", X$, REG&)
          End If
        End If
        '
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) + 0.005 Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If Abs(SALDO#) > 0 Then
            FF% = CVI(Mid$(X$, 9, 2))
            Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
            LINDICE.AddItem Z$
        End If
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To LINDICE.ListCount
        LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste de Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota de Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Descuento Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota de Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste de Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                SUMAFAC# = SUMAFAC# + IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROL& = CVS(Mid$(X$, 5, 4))
                FEMI = CVI(Mid$(X$, 19, 2))
                FEVE = CVI(Mid$(X$, 9, 2))
                '
1234            FEVEX$ = FECHATEX$(FEVE)
                If Val(Mid$(FEVEX$, 4, 2)) = 2 Then
                  If Val(Mid$(FEVEX$, 1, 2)) > 28 Then
                    FEVE = DIPOST(FEVE)
                    GoTo 1234
                  End If
                End If
                If CVI(Mid$(X$, 9, 2)) <> FEVE Then
                  FEVEI% = FEVE
                  Call REPLA(X$, MKI$(FEVEI%), 9, 2)
                  Call GRAREG("SVD202", X$, REG&)
                End If
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                '
If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
  MONEMX% = Asc(Mid$(X$, 31, 1))
  If MONEMX% > 2 Then
    MONEMX% = 0
    Call REPLA(X$, Chr$(0), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 0 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          MONEMX% = Asc(Mid$(X$, 31, 1))
          GoTo 15
        End If
      End If
    Next KUL&
15 If MONEMX% <> 2 Then MONEMX% = 1
    Call REPLA(X$, Chr$(MONEMX%), 31, 1)
    Call GRAREG("SVD202", X$, REG&)
  End If
  If MONEMX% = 2 Then
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
        If Asc(Mid$(XXC$, 37, 1)) = TI% Then
          SALTEX$ = "u$s" + FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F11.2")
          GoTo 20
        End If
      End If
    Next KUL&
  End If
End If
'
20              If TRIM$(SALTEX$) <> "" Then
                  Z$ = Space$(80)
                  Call REPLA(Z$, TIPOS$, 1, 16)
                  Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 17, 8)
                  Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
                  Call REPLA(Z$, IMPTEX$, 37, 14)
                  Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
                  Call REPLA(Z$, SALTEX$, 63, 14)
                  '
                  List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label28.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label28.Refresh
    Label25.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label25.Refresh
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       ATENTI% = 1
       TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
       If Visible = True Then
           TTXX$ = TTXX$ + " Elegida."
         Else
           TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NC%)) + "'"
       End If
       TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
       '
       If COMALTER%(TTXX$) = 2 Then
           Call REAPLICRE(NC%)
           GoTo 50
         Else
           ATENTI% = 0
           Call CIERRARCH("*.*")
           Call FINAL("")
       End If
    End If
    '
    If Abs(SARRAS# - SALDCLI#(NC%)) >= 0.005 Then
        RGXC& = REGICLI&(NC%)
        XC$ = REGLEIDO$("DEUDOR2", RGXC&)
        Call REPLA(XC$, MKD$(SARRAS#), 109, 8)
        Call GRAREG("DEUDOR2", XC$, RGXC&)
        LTEXT37.Caption = CSTRING$(MKD$(SARRAS#), 4, 11, 2, 2)
        Refresh
    End If
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
100 Screen.MousePointer = 11
    '
    SUMARRAS# = 0
    If List3.ListCount > 0 Then
      SUMARRAS# = Val(Mid$(List3.List(List3.ListCount - 1), 63, 14))
    End If
    SUMADEUP# = Val(Label28.Caption)
    CADOCPEN% = List2.ListCount
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    ' para no enquilombar las cuentas que mas o menos estan
    If Abs(SUMARRAS# - SUMADEUP#) < 0.005 Then
      If NUDOCU% >= CADOCPEN% Then
        GoTo 999
      End If
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
             SG% = 1: If TIPO% >= 4 Then If TIPO% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          ARDEX$ = ARCHIDEU$ + "/" + TRIM$(CSTRING$(MKD$(IORIG#), 4, 12, 2, 2))
          Call GRADEU(LU$, ARDEX$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub


Private Sub Text36_Change()
   MODIFIC%(0) = 1
End Sub
Sub TEXT36_LOSTFOCUS()
    Text36 = FORMATNUM(Val(Text36), "F6.1")
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

Sub CARGOTR(NUOTR$) 'CARGA LA FICHA DE UNA OT DETERMINADA
   '
   If NUOTR$ = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM ORDETRAB"
   Set OTTMP = OTRABAJO.OpenRecordset(SQLX$, 4)
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
     NCLIE% = !NCLIE_oTRAB
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
     Text12 = FECHATEX(CVINT(!FeOFer_Otrab))
     Text17 = !Especi_oTRAB
     Text36 = !CanPed_Otrab
     Text38 = !preven_Otrab
     Text38 = FORMATNUM$(Val(Text38), "F12.2")
     '
     Text2 = !Lentrega_Otrab
     COTEXT(5) = !Vended_Otrab
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
     '
   End With
   Call SUMAHS
   CAMBIOT% = 0
   '
End Sub
Sub GENEINDIOTRAB() 'GENERA EL INDICE DE ORDENES DE TRABAJO
   'If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABREORTRA
      Set ORDETRATECO = OTRABAJO.OpenRecordset("SELECT NRO_oTRAB, REFE_oTRAB,FEMI_oTRAB FROM ORDETRAB  ORDER BY REFE_OTRAB ASC;", dbOpenDynaset, dbReadOnly)
      On Error Resume Next
      ORDETRATECO.MoveLast
      If Err Then Exit Sub
      '
      CAREC% = ORDETRATECO.RecordCount
      If CAREC% < 1 Then Exit Sub
      '
      FIN& = CAREC%
      'REBLA$ = REGBLAN$("DEUDOR1")
      REBLB$ = REGBLAN$("INDIOTRA")
      JJ& = 0
      With ORDETRATECO
        .MoveFirst
        '
10      Call TRACE(20, JJ&, FIN&)
        'EL1$ = REB
       ' Call REPLA(EL1$, MKI$(!Nume_Cli), 1, 2)
        'Call REPLA(EL1$, !Raso_Cli, 3, 30)
        JJ& = JJ& + 1
        'Call GRAREG("DEUDOR1", EL1$, JJ&)
        '
        EL2$ = REBLB$
        Call REPLA(EL2$, !Nro_Otrab, 1, 16)
        Call REPLA(EL2$, !Refe_Otrab, 17, 48)
        'Call REPLA(EL2$, MKI$(!Nume_Cli), 63, 2)
        Call GRAREG("INDIOTRA", EL2$, JJ&)
        '
        If JJ& < CAREC% Then
          ORDETRATECO.MoveNext
          GoTo 10
        End If
      End With
      '
      ORDETRATECO.Close
      'Call SETULTREG("DEUDOR1", JJ&)
      'Call CIERRARCH("DEUDOR1")
      Call FILESORT("INDIOTRA", "", 1, 1, "ASC")
      'Call FRESHECHO("DEUDOR1")
      Call FRESHECHO("INDIOTRA")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
   'End If
End Sub
Sub PRIORVEN()
    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
'    Call SETULTREG("!OBSERVOF", 0)
'
'    FORIPRE$ = CONTROL$("", "FOROTRAB")
'ANTEXTO = 52
'    Call FRATEXTO(Text35, 52)
'    For KKU& = 1 To CARETEX%
'      Call REGAPP("!OBSERVOF", RETEX$(KKU&))
'    Next KKU&
'    Call CIERRARCH("!OBSERVOF")
    '
    PRIPLANOI% = 0
    PRIESPEC% = 0
    PRIHRINS% = 0
    '
    ' IMPRESION DE LA Especificación de Vta.
    '
    FORIPRE$ = CONTROL$("", "ESVE-VTH")
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
      CODPARAM$(21) = "NRO-PASE"
      
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
      DOCPARAM(18) = Text34
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
      Call PRINTDOC("ESVE-VTH")                   ' Orden de TRABAJO
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

