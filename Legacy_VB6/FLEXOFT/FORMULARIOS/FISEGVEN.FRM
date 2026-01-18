VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FISEGVEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento Comercial de Ventas"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11775
   Icon            =   "FISEGVEN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11775
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404000&
      Height          =   8415
      Left            =   10815
      ScaleHeight     =   8355
      ScaleWidth      =   1110
      TabIndex        =   94
      Top             =   0
      Width           =   1170
      Begin VB.Frame Frame6 
         BackColor       =   &H00404000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   80
         TabIndex        =   105
         Top             =   5565
         Width           =   780
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
            Height          =   580
            Left            =   105
            Picture         =   "FISEGVEN.frx":1272
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Padron de Vendedores"
            Top             =   945
            Width           =   600
         End
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   105
            Picture         =   "FISEGVEN.frx":157C
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   315
            Width           =   600
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00404000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   80
         TabIndex        =   102
         Top             =   3570
         Width           =   780
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   105
            Picture         =   "FISEGVEN.frx":19BE
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Imprime Ficha del Contacto"
            Top             =   315
            Width           =   600
         End
         Begin VB.CommandButton Command18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   580
            Left            =   105
            Picture         =   "FISEGVEN.frx":2028
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Listados por Impresora"
            Top             =   945
            Width           =   600
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   100
         Top             =   7455
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   101
            Top             =   0
            Width           =   12000
         End
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
         Height          =   600
         Left            =   105
         Picture         =   "FISEGVEN.frx":2332
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1470
         Width           =   750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Sort"
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
         Left            =   105
         Picture         =   "FISEGVEN.frx":263C
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Ordenamiento Alfabético"
         Top             =   2730
         Width           =   750
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
         Height          =   600
         Left            =   105
         Picture         =   "FISEGVEN.frx":2946
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   750
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
         Height          =   600
         Left            =   105
         Picture         =   "FISEGVEN.frx":2A90
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Alta de Nueva Ficha de Seguimiento"
         Top             =   840
         Width           =   750
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Erase"
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
         Left            =   105
         Picture         =   "FISEGVEN.frx":2BDA
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Elimina Contacto Activo de la Base"
         Top             =   2100
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "FISEGVEN.frx":2EE4
         Top             =   7785
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Misc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1950
      Left            =   9660
      TabIndex        =   77
      Top             =   6300
      Width           =   1065
      Begin VB.CommandButton Command17 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FISEGVEN.frx":4E06
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Documentos Conexos"
         Top             =   1155
         Width           =   850
      End
      Begin VB.CommandButton Command11 
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
         Height          =   740
         Left            =   105
         Picture         =   "FISEGVEN.frx":5110
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Abrir Anotador de Novedades"
         Top             =   315
         Width           =   850
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4425
      Left            =   105
      TabIndex        =   34
      Top             =   105
      Width           =   10620
      Begin VB.CommandButton Command8 
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
         Left            =   10220
         TabIndex        =   108
         ToolTipText     =   "Selecciona Vendedor"
         Top             =   2940
         Width           =   175
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   9345
         TabIndex        =   93
         Text            =   " "
         Top             =   1050
         Width           =   1065
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
         Left            =   8925
         TabIndex        =   89
         Top             =   1050
         Width           =   150
      End
      Begin VB.TextBox Text13 
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
         Left            =   1155
         TabIndex        =   87
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   1050
         Width           =   6630
      End
      Begin VB.CommandButton Command23 
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
         Height          =   260
         Left            =   8820
         TabIndex        =   84
         Top             =   4070
         Width           =   790
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   9620
         TabIndex        =   83
         Top             =   4070
         Width           =   790
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8820
         Sorted          =   -1  'True
         TabIndex        =   81
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3675
         Width           =   1590
      End
      Begin VB.CommandButton Command19 
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
         Left            =   7500
         Picture         =   "FISEGVEN.frx":541A
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Marca / Quita Prioridad"
         Top             =   420
         Width           =   285
      End
      Begin VB.CommandButton Command13 
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
         Left            =   9450
         Picture         =   "FISEGVEN.frx":585C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Selección Contactos Programados"
         Top             =   420
         Width           =   480
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3360
         Width           =   7260
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
         Height          =   280
         Left            =   9660
         MaxLength       =   3
         TabIndex        =   16
         Text            =   " "
         Top             =   2940
         Width           =   600
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
         Height          =   280
         Left            =   1155
         TabIndex        =   14
         Text            =   " "
         Top             =   2940
         Width           =   2115
      End
      Begin VB.TextBox Text2 
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
         Left            =   4830
         TabIndex        =   15
         Text            =   " "
         Top             =   2940
         Width           =   3585
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9930
         Picture         =   "FISEGVEN.frx":5C9E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Búsqueda Secuencial por Texto"
         Top             =   420
         Width           =   480
      End
      Begin VB.TextBox Text5 
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
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   735
         Width           =   6630
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nro"
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
         Left            =   7875
         TabIndex        =   53
         Top             =   315
         Width           =   1520
         Begin VB.CommandButton Command27 
            Caption         =   ">"
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
            Left            =   1280
            Picture         =   "FISEGVEN.frx":65A0
            TabIndex        =   86
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   160
         End
         Begin VB.CommandButton Command26 
            Caption         =   "<"
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
            Left            =   1120
            Picture         =   "FISEGVEN.frx":68AA
            TabIndex        =   85
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   160
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
            Left            =   885
            TabIndex        =   0
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   175
         End
         Begin VB.Label Label26 
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
            Left            =   150
            TabIndex        =   54
            Top             =   260
            Width           =   750
         End
      End
      Begin VB.TextBox Text4 
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
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   420
         Width           =   6420
      End
      Begin VB.TextBox Text26 
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
         Left            =   7875
         TabIndex        =   11
         Text            =   " "
         Top             =   1890
         Width           =   2150
      End
      Begin VB.TextBox Text24 
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
         Left            =   7875
         TabIndex        =   13
         Text            =   " "
         Top             =   2520
         Width           =   2150
      End
      Begin VB.TextBox Text23 
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
         Left            =   1155
         TabIndex        =   9
         Text            =   " "
         Top             =   2520
         Width           =   5895
      End
      Begin VB.TextBox Text22 
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
         Left            =   3675
         TabIndex        =   7
         Text            =   " "
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox Text21 
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
         Left            =   1155
         TabIndex        =   6
         Top             =   1785
         Width           =   1485
      End
      Begin VB.TextBox Text20 
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
         Left            =   1155
         TabIndex        =   5
         Text            =   " "
         Top             =   1470
         Width           =   5895
      End
      Begin VB.TextBox Text19 
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
         Left            =   1155
         TabIndex        =   8
         Text            =   " "
         Top             =   2205
         Width           =   5895
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
         Height          =   280
         Left            =   7875
         TabIndex        =   10
         Text            =   " "
         Top             =   1470
         Width           =   2535
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
         Height          =   280
         Left            =   7875
         TabIndex        =   12
         Text            =   " "
         Top             =   2205
         Width           =   2535
      End
      Begin VB.Label Label40 
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
         Height          =   260
         Left            =   8610
         TabIndex        =   109
         Top             =   3255
         Width           =   1780
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "M:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9135
         TabIndex        =   92
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8250
         TabIndex        =   91
         Top             =   1050
         Width           =   680
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Baja:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7875
         TabIndex        =   90
         ToolTipText     =   "Doble-Click para Re-Activar"
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Web-Site:"
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
         Left            =   105
         TabIndex        =   88
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Gr:"
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
         Left            =   7845
         TabIndex        =   82
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3660
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   1155
         X2              =   8400
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   105
         TabIndex        =   59
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   8540
         TabIndex        =   58
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ramo:"
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
         Left            =   105
         TabIndex        =   57
         Top             =   3045
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Princ.Producto"
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
         Left            =   2940
         TabIndex        =   56
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         Left            =   105
         TabIndex        =   55
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
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
         Left            =   105
         TabIndex        =   52
         Top             =   525
         Width           =   960
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   10080
         Picture         =   "FISEGVEN.frx":6BB4
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10080
         Picture         =   "FISEGVEN.frx":6EBE
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Nº:"
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
         Left            =   105
         TabIndex        =   42
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad:"
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
         Left            =   2625
         TabIndex        =   41
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "País:"
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
         Left            =   105
         TabIndex        =   40
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Telef:"
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
         Left            =   6825
         TabIndex        =   39
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "C.Post:"
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
         Left            =   105
         TabIndex        =   38
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
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
         Left            =   6825
         TabIndex        =   37
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia:"
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
         Left            =   105
         TabIndex        =   36
         Top             =   2310
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Origen:"
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
         Left            =   6825
         TabIndex        =   35
         Top             =   1575
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1950
      Left            =   105
      TabIndex        =   65
      Top             =   6300
      Width           =   9460
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
         Height          =   280
         Left            =   8295
         TabIndex        =   75
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
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
         Left            =   7350
         TabIndex        =   76
         Top             =   1470
         Width           =   1000
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   9360
         TabIndex        =   66
         Top             =   315
         Width           =   9420
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle                                                    Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   68
            Top             =   75
            Width           =   9360
         End
         Begin VB.Label Label13 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   67
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8670
         Picture         =   "FISEGVEN.frx":71C8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1140
         Width           =   750
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   735
         Width           =   7155
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Label35"
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
         Left            =   7245
         TabIndex        =   74
         Top             =   840
         Width           =   4005
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   80
         TabIndex        =   73
         Top             =   800
         Width           =   1065
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Próx.Contacto:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7245
         TabIndex        =   72
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Line Line4 
         X1              =   7140
         X2              =   9450
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Label Label32 
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
         Height          =   1170
         Left            =   7140
         TabIndex        =   71
         Top             =   735
         Width           =   2310
      End
      Begin VB.Label Label31 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   70
         Top             =   735
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1920
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Line Line7 
         X1              =   9450
         X2              =   9450
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   69
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   11665
         Y1              =   1900
         Y2              =   1900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contactos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   105
      TabIndex        =   60
      Top             =   4560
      Width           =   10620
      Begin VB.TextBox Text10 
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
         Index           =   2
         Left            =   9135
         TabIndex        =   29
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox Text8 
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
         Index           =   2
         Left            =   7140
         TabIndex        =   28
         Top             =   1260
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Index           =   2
         Left            =   4830
         TabIndex        =   27
         Top             =   1260
         Width           =   2115
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
         Height          =   280
         Index           =   2
         Left            =   1155
         TabIndex        =   26
         Top             =   1260
         Width           =   3480
      End
      Begin VB.TextBox Text10 
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
         Index           =   1
         Left            =   9135
         TabIndex        =   25
         Top             =   945
         Width           =   1275
      End
      Begin VB.TextBox Text8 
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
         Index           =   1
         Left            =   7140
         TabIndex        =   24
         Top             =   945
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Index           =   1
         Left            =   4830
         TabIndex        =   23
         Top             =   945
         Width           =   2115
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
         Height          =   280
         Index           =   1
         Left            =   1155
         TabIndex        =   22
         Top             =   945
         Width           =   3480
      End
      Begin VB.TextBox Text10 
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
         Index           =   0
         Left            =   9135
         TabIndex        =   21
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox Text8 
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
         Index           =   0
         Left            =   7140
         TabIndex        =   20
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Index           =   0
         Left            =   4830
         TabIndex        =   19
         Top             =   630
         Width           =   2115
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
         Height          =   280
         Index           =   0
         Left            =   1155
         TabIndex        =   18
         Top             =   630
         Width           =   3480
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Directo / Interno:"
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
         Left            =   8085
         TabIndex        =   64
         Top             =   350
         Width           =   2010
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Sector:"
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
         Left            =   6510
         TabIndex        =   63
         Top             =   350
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo / Función:"
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
         Left            =   4305
         TabIndex        =   62
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Persona de Contacto:"
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
         Left            =   210
         TabIndex        =   61
         Top             =   350
         Width           =   1905
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10815
      Sorted          =   -1  'True
      TabIndex        =   51
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      Height          =   2535
      Left            =   2835
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   43
      Top             =   1890
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
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
         Height          =   240
         Left            =   3885
         TabIndex        =   50
         Top             =   -10
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1260
         TabIndex        =   49
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levante el Receptor !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   420
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   210
         TabIndex        =   47
         Top             =   1365
         Width           =   3795
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "45241284"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   420
         TabIndex        =   46
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Llamando a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   45
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         Caption         =   "Monitor de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   4110
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   33
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   32
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FISEGVEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelflag

Private Sub Command11_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command11.Enabled = False
   NPX& = Val(Label26)
   '
   If NPX& > 0 Then
     ANOTASEGVE.Label26 = Label26
     ANOTASEGVE.Label3 = Text4
     ANOTASEGVE.Label7 = Text6(0)
     ANOTASEGVE.Label8 = Text6(1)
     ANOTASEGVE.Label11 = Text1
     ANOTASEGVE.Label13 = Text2
     ANOTASEGVE.Label14 = TRIM$(Text3)
     ANOTASEGVE.Text12 = Text12
     '
     ANOTASEGVE.List1.Clear
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       For KU% = 5 To 92
         If Asc(Mid$(XX$, KU%)) < 32 Then
           Mid$(XX$, KU%, 1) = " "
         End If
      Next KU%
      ANOTASEGVE.List1.AddItem Mid$(XX$, 5, 92)
     Next U&
     '
     TPI = ANOTASEGVE.List1.ListCount - 12
     If TPI < 0 Then TPI = 0
     ANOTASEGVE.List1.TopIndex = TPI
     '
     ANOTASEGVE.Show 1
   End If
   '
   Command11.Enabled = True
   '
End Sub

Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.Text = VDEV$
    End If
End Sub

Private Sub Command16_Click()
   If TRIM$(FORSELEC.Text1) = "" Then
     FORSELEC.ABUSQUE = TRIM$(Label26)
   End If
   Call SELECHO("BSEGCOM")
   VDEVU$ = VALACT1$("BSEGCOM")
   If VDEVU$ <> "" Then
     Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command17_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command17.Enabled = False
   NPX& = Val(Label26)
   '
   If NPX& > 0 Then
     DOCONECVE.Label26 = Label26
     DOCONECVE.Label5 = Text4
     DOCONECVE.Label7 = Text6(0)
     DOCONECVE.Label8 = Text6(1)
     DOCONECVE.Label4 = Text1
     DOCONECVE.Label13 = Text2
     DOCONECVE.Label14 = TRIM$(Text3)
     '
     DOCONECVE.List1.Clear
     NPXI% = NPX&
     RFF$ = RECFILT$("DOCONEC", 1, MKI$(NPXI%))
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("DOCONEC", RENOTA&)
       YY$ = Space$(256)
       Call REPLA(YY$, Mid$(XX$, 3, 56), 1, 56)
       Call REPLA(YY$, Mid$(XX$, 129, 56), 61, 56)
       Call REPLA(YY$, TRIM$(Str$(RENOTA&)), 129, 6)
       DOCONECVE.List1.AddItem YY$
     Next U&
     '
     DOCONECVE.Show 1
   End If
   '
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   OPX% = COMALTER%("Opciones de Listados:\\Padrón de Contactos\Actividad por Vendedor")
   If OPX% = 1 Then
       Call FINAL("*BSEGCOM")
     ElseIf OPX% = 2 Then
       Call LIACVENDE
   End If
   Command18.Enabled = True
   '
End Sub

Sub LIACVENDE()
   Dim CLISELE%()
   HOII% = HOY(HOS$)
   VDEF$ = MKI$(COMESACT%) + MKI$(HOII%) + Chr$(0)
10 VDEVU$ = OBJPLA$("DESHASFECHVEN", VDEF$)
   If VDEVU$ <> "" Then
     VDEF$ = VDEVU$
     DES% = CVI(Left$(VDEVU$, 2))
     HAS% = CVI(Mid$(VDEVU$, 3, 2))
     If HAS% < DES% Then
        Call MENSERR(24, "Rango de Fechas no Válido")
        GoTo 10
     End If
     NVENDE% = Asc(Mid$(VDEVU$, 5, 1))
     If NVENDE% < 1 Or ECOARCH$("VENDEDOR", Chr$(NVENDE%)) = "" Then
        Call MENSERR(24, "Falta Vendedor")
        GoTo 10
     End If
   End If
   '
   Screen.MousePointer = 11
   DENOVE$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(NVENDE%)))
   FEX = DES%: DESX$ = FECHATEX$(FEX)
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   '
   FIN& = ULTREG&("BSEGCOM")
   JJ& = 0
   For KKU& = 1 To FIN&
     Call TRACE(20, KKU&, FIN&)
     X$ = REGLEIDO$("BSEGCOM", KKU&)
     If XVALO(Mid$(X$, 369, 8)) = NVENDE% Then
       CACLI% = CACLI% + 1
       ReDim Preserve CLISELE%(CACLI%)
       CLISELE%(CACLI%) = CVI(Left$(X$, 2))
     End If
   Next KKU&
   '
   Call BLOQARCH("LISEGVEN")
   REBLA$ = REGBLAN$("LISEGVEN")
   For KKI% = 1 To CACLI%
     CLISE& = CLISELE%(KKI%)
     DENCLI$ = ECOARCH$("BSEGCOM", MKI$(CLISELE%(KKI%)))
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(CLISE&))
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       If TRIM$(Mid$(XX$, 5, 8)) <> "" Then
         If TRIM$(Mid$(XX$, 5, 8)) <> "========" Then
           FECHAI% = FECHANUM(Mid$(XX$, 5, 8))
           If FECHAI% >= DES% Then
             If FECHAI% <= HAS% Then
               For U2& = U& To Len(RFF$) Step 4
                 RENOTA2& = CVS(Mid$(RFF$, U2&, 4))
                 YY$ = REGLEIDO$("ANOTASEG", RENOTA2&)
                 ZZ$ = REBLA$
                 Call REPLA(ZZ$, MKI$(FECHAI%), 1, 2)
                 Call REPLA(ZZ$, MKI$(FECHAI%), 49, 2)
                 Call REPLA(ZZ$, DENCLI$, 3, 46)
                 Call REPLA(ZZ$, Mid$(YY$, 15, 56), 51, 56)
                 JJ& = JJ& + 1
                 Call GRAREG("LISEGVEN", ZZ$, JJ&)
                 If U2& > U& Then
                   If TRIM$(Mid$(YY$, 5, 8)) <> "" Then
                     GoTo 26
                   End If
                 End If
               Next U2&
             End If
           End If
         End If
       End If
26   Next U&
   Next KKI%
   '
   Call SETULTREG("LISEGVEN", JJ&)
   Call BLOQARCH("LISEGVEN")
   Call HEADERS("LISEGVEN", "")
   Call HEADERS("LISEGVEN", "Vendedor: " + DENOVE$ + " - Desde " + DESX$ + " hasta " + HASX$)
   Screen.MousePointer = 1
   '
   OPX% = COMALTER%("Elija Ordenamiento:\\Por Fecha\Por Contacto")
   If OPX% = 1 Then
       Call FILESORT("LISEGVEN", "", 1, 2)
       Call FINAL("*LISEGVE1")
     ElseIf OPX% = 2 Then
       Call FILESORT("LISEGVEN", "", 2, 3)
       Call FINAL("*LISEGVE2")
   End If
   Call LIBARCH("LISEGVEN")
   '
End Sub
'
Private Sub Command19_Click()
   RCONTA$ = Text4 + Space$(64)
   If Mid$(RCONTA$, 60, 3) = "!!!" Then
       Mid$(RCONTA$, 60, 3) = "   "
     Else
       Mid$(RCONTA$, 60, 3) = "!!!"
   End If
   Text4 = TRIM$(RCONTA$)
End Sub

Private Sub Command2_Click()
   '
   Static ARGUME$
   Command2.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
10 ARGUME$ = TRIM$(UCase$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGUME$)))
   '
   If ARGUME$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     For U& = 1 To ULTREG&("BSEGCOM")
       XXX$ = REGLEIDO$("BSEGCOM", U&)
       If InStr(UCase$(Mid$(XXX$, 3)), ARGUME$) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("!SECONTAC", XXX$, JJ&)
       End If
     Next U&
     '
     Call SETULTREG("!SECONTAC", JJ&)
     Call CIERRARCH("!SECONTAC")
     Call BAJALDISCO
     Call FRESHECHO("!SECONTAC")
     '
20   If JJ& > 0 Then
       Call SELECHO("!SECONTAC")
       NPX& = Val(VALACT1$("!SECONTAC"))
       If NPX& > 0 Then
         Label26 = Str$(NPX&)
       End If
     End If
     '
   End If
   Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   If List2.ListIndex >= 0 Then
     List2.RemoveItem (List2.ListIndex)
   End If
End Sub

Private Sub Command23_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   XX$ = REGSEL$("TAGRUCON")
   If Len(XX$) > 4 Then
     RAGRE% = Val(CVI(Left$(XX$, 2)))
     TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 24) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
     For U& = 1 To List2.ListCount
       If Val(Mid$(List2.List(U& - 1), 72, 6)) = RAGRE% Then Exit Sub
     Next U&
     List2.AddItem TAGRE$
   End If
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If COMALTER%("Realmente Desea Eliminar PERMANENTEMENTE\la Ficha Activa ?\\No, Cancelar\Si, Eliminar") = 2 Then
5    Screen.MousePointer = 11
     IFICHA% = Val(Label26)
     '
     ABASE$ = "BSEGCOM": GoSub BORRABASE
     ABASE$ = "DOCONEC": GoSub BORRABASE
     ABASE$ = "NOTASEG": GoSub BORRABASE
     '
     ABASE$ = "ANOTASEG": GoSub BORRABASLAR
     '
     Label26 = ""
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   Exit Sub
   '
   '
BORRABASE:
   JJ& = 0
   For U& = 1 To ULTREG&(ABASE$)
     XX$ = REGLEIDO$(ABASE$, U&)
     IFICHB% = CVI(Left$(XX$, 2))
     If IFICHB% <> IFICHA% Then
       If IFICHB% > IFICHA% Then IFICHB% = IFICHB% - 1
       Call REPLA(XX$, MKI$(IFICHB%), 1, 2)
       JJ& = JJ& + 1
       Call GRAREG(ABASE$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(ABASE$, JJ&)
   Call CIERRARCH(ABASE$)
Return
   
BORRABASLAR:
   JJ& = 0
   For U& = 1 To ULTREG&(ABASE$)
     XX$ = REGLEIDO$(ABASE$, U&)
     IFICHC& = CVS(Left$(XX$, 4))
     If IFICHC& <> IFICHA% Then
       If IFICHC& > IFICHA% Then IFICHC& = IFICHC& - 1
       Call REPLA(XX$, MKS$(IFICHC&), 1, 4)
       JJ& = JJ& + 1
       Call GRAREG(ABASE$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(ABASE$, JJ&)
   Call CIERRARCH(ABASE$)
   '
Return
'
'
End Sub

Private Sub Command26_Click()
   UKU& = REGNUM&("BSEGCOM")
   If UKU& > 1 Then
     XX$ = REGLEIDO$("BSEGCOM", UKU& - 1)
     Label26 = TRIM$(Str$(CVI(Left$(XX$, 2))))
   End If
End Sub

Private Sub Command27_Click()
   UKU& = REGNUM&("BSEGCOM")
   If UKU& < ULTREG&("BSEGCOM") Then
     XX$ = REGLEIDO$("BSEGCOM", UKU& + 1)
     Label26 = TRIM$(Str$(CVI(Left$(XX$, 2))))
   End If
End Sub

Private Sub Command28_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Label38 = VDEV$
    End If
End Sub



Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Label35 = NOMTER
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))

   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   MSComm1.InputLen = 0
   'If Screen.Width < 12000 Then
   '   Hide
   '   FICONTAX.Show
   'End If
   '
End Sub

Private Sub Command1_Click()
   '
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command1.Enabled = False
   If Val(Label26) > 0 Then
     Call GRAFICON
   End If
   Command1.Enabled = True
   '
End Sub

Private Sub Command10_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command10.Enabled = False
   Call EXPORAGE
   Command10.Enabled = True
End Sub

Private Sub Command12_Click()
   '
   Command12.Enabled = False
   IMPORCONTAC.Show 1
   'Call IMPORAGE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static FIDATX1#, FIDATX2#, UCONTA$
   '
   Command13.Enabled = False
   If FIDATIM#("BSEGCOM") = FIDAT1X# Then
     If FIDATIM#("!SECONTAC") = FIDATX2# Then
       JJ& = ULTREG&("!SECONTAC")
       GoTo 20
     End If
   End If
   '
   FIDATX1# = FIDATIM#("BSEGCOM")
   JJ& = 0
   HOII% = HOY(HOS$)
   FINIPRO% = FECHANUM("30/04/06")  ' PARA EXCLUIR CONTACTOS VIEJOS
   For U& = 1 To ULTREG&("BSEGCOM")
     XXX$ = REGLEIDO$("BSEGCOM", U&)
     If CVI(Mid$(XXX$, 1023, 2)) <= HOII% Then
       If CVI(Mid$(XXX$, 1023, 2)) > FINIPRO% Then
         JJ& = JJ& + 1
         Call GRAREG("!SECONTAC", XXX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("!SECONTAC", JJ&)
   Call CIERRARCH("!SECONTAC")
   Call BAJALDISCO
   FIDATX2# = FIDATIM#("!SECONTAC")
   Call FRESHECHO("!SECONTAC")
   '
20 If JJ& > 0 Then
     FORSELEC.ABUSQUE = UCONTA$
     Call SELECHO("!SECONTAC")
     UCONTA$ = VALACT1$("!SECONTAC")
     NPX& = Val(UCONTA$)
     If NPX& > 0 Then
       Label26 = Str$(NPX&)
     End If
   End If
   '
   Command13.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPX% = COMALTER%("Opciones de Configuración:\\Conf.General\Particular")
   If OPX% = 1 Then
        Call CONECRUN("FLSETUP", "")
      ElseIf OPX% = 2 Then
        SETUPSGC.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command21_Click()
   Label26 = 1 + ULTREG&("BSEGCOM")
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = HOS$
   On Error Resume Next
   Text4.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command25_Click()
   ANCHAC = Width
   VISIAC = Picture1.Visible
   Image2.Visible = False
   Picture1.Visible = False
   Width = 10800
   On Error Resume Next
   PrintForm
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Image2.Visible = True
   Picture1.Visible = VISIAC
   Width = ANCHAC
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command4_Click()
   cancelflag = 1
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   HOII% = HOY(HOS$)
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
     If TRIM$(Text11) <> "" Then
       Call FRATEXTO(Text11, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, NOMTER$, 69, 24)
         End If
         Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTASEG", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTASEG")
       Text11 = ""
       '
       NPX1% = NPX&
       RFF$ = RECFILT$("BSEGCOM", 1, MKI$(NPX1%))
       If Len(RFF$) >= 4 Then
         NPY& = CVS(Left$(RFF$, 4))
         If NPY& <= ULTREG&("BSEGCOM") Then
           HOII% = HOY(HOS$)
           FEPROCON% = FECHANUM(Text12)
           If FEPROCON% < HOII% Then FEPROCON% = HOII%
           '
           XXX$ = REGLEIDO$("BSEGCOM", NPY&)
           If CVI(Left$(XXX$, 2)) = NPX1% Then
             Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
             Call GRAREG("BSEGCOM", XXX$, NPY&)
             Call CIERRARCH("BSEGCOM")
           End If
         End If
       End If
       '
     End If
   End If
   '
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   cancelflag = 1
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   Call TRALOCAL("VENDEDOR", "")
10 VTB% = VENTABU%("VENDEDOR")
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      Dim CPAX%(255)
      For U& = 1 To ULTREG&("!VENDEDOR")
         X$ = REGLEIDO$("!CONPAG", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Or COCO% > 255 Then
            Call MENSERR(24, "Vendedor Número '" & COCO% & "' no Válido.")
            GoTo 10
         End If
         For V& = 1 To U& - 1
            If COCO% = CPAX%(V&) Then
               Call MENSERR(24, "Vendedor Número '" + TRIM$(Str$(COCO%)) + "' Repetido.")
               GoTo 10
            End If
         Next V&
         CPAX%(U&) = COCO%
         '
         DECO$ = TRIM$(Mid$(X$, 2, 35))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Apellido y Nombre del\Vendedor Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("VENDEDOR", "")
      Call FRESHECHO("VENDEDOR")
   End If
   '
99 Command7.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command8_Click()
   '
   Call SELECHO("VENDEDOR")
   NVENDE% = XVALO(VALACT1$("VENDEDOR"))
   If NVENDE% > 0 Then
     Text3 = TRIM$(Str$(NVENDE%))
   End If
   '
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   Call SORCONTAC
   Command9.Enabled = True
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image4_Click()
   Image4.Enabled = False
   Call Dial(TRIM$(Text26.Text), Name)
   Image4.Enabled = True
End Sub

Private Sub Image5_Click()
   Image5.Enabled = False
   Call Dial(TRIM$(Text24.Text), Name)
   Image5.Enabled = True
End Sub

Private Sub Label16_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Label26_Change()
   '
   Call CARFICHA
   '
End Sub

Private Sub Label37_DblClick()
   Label38 = ""
   Text14 = ""
End Sub

Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.Text, 25, 48))
End Sub

Private Sub List2_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Text1_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text11_DblClick()
   Text11.Enabled = False
   Call Dial(TRIM$(Text11.Text), Name)
   Text11.Enabled = True
End Sub

Private Sub Text11_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.Text = VDEV$
    End If
End Sub

Private Sub Text12_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text17_DblClick()
   Text17.Enabled = False
   Call Dial(TRIM$(Text17.Text), Name)
   Text17.Enabled = True
End Sub

Private Sub Text17_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text18_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text19_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text20_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text21_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text22_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text23_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text24_DblClick()
   Text24.Enabled = False
   Call Dial(TRIM$(Text24.Text), Name)
   Text24.Enabled = True
End Sub

Private Sub Text24_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text26_DblClick()
   Text26.Enabled = False
   Call Dial(TRIM$(Text26.Text), Name)
   Text26.Enabled = True
End Sub

'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   por Carl Franklin
'
'   Actualizado por Anton de Jong
'
'   Muestra cómo marcar números telefónicos con un módem.
'
'   Para que este programa funcione, el teléfono y el
'   módem deben estar conectados a la misma línea.
'--------------------------------------------------------
'Option Explicit

'DefInt A-Z

' Este indicador se establece cuando el usuario elige Cancelar.

Sub Dial(NULLAMA$, FORNAME$)
   '
   Dim DialString$, FromModem$, dummy
    ' AT es el comando compatible Hayes ATTENTION, y es necesario para enviar comandos al módem.
    ' DT significa "marcar con tonos". El comando Dial utiliza tonos, en lugar de pulsos (DP =
    ' marcar con pulsos). Number$ es el número de teléfono a marcar.
    ' El punto y coma indica al módem que vuelva al modo de comando después de marcar
    ' (importante). Es necesario un retorno de carro, vbCr, al enviar comandos al módem.
    '
   '
   For ki% = 1 To Forms.Count
     If Forms(ki% - 1).Name = FORNAME$ Then GoTo 5
   Next ki%
   Exit Sub
   '
5  Number$ = ""
   INTERN$ = ""
   TOMALIN$ = TRIM$(CONTROL("AGETELEF", "TOMALIN"))
   '
   PPXX% = InStr(NULLAMA$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NULLAMA$)
   INTERN$ = TRIM$(Mid$(NULLAMA$, PPXX% + 1))
   NULLAMA$ = TRIM$(Left$(NULLAMA$, PPXX% - 1))
   '
   Number$ = TOMALIN$
   For KKI% = 1 To Len(NULLAMA$)
     CARA% = Asc(Mid$(NULLAMA$, KKI%, 1))
     If CARA% = 44 Or CARA% >= 48 Then
       If CARA% <= 57 Then
         Number$ = Number$ + Chr$(CARA%)
       End If
     End If
   Next KKI%
   If Number$ = "" Then Exit Sub
   '
10 Label28 = Number$
   Label29 = ""
   Picture1.Visible = True
   Picture1.Refresh
   cancelflag = 0
   '
   DialString$ = "ATDT" + Number$ + ";" + vbCr
    '
    ' Configuración del puerto de comunicaciones.
    ' Suponiendo que hay conectado un mouse a COM1, se establece 2 en CommPort
   SETUPSTRI$ = CONTROL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   COPORT = Val(Left$(SETUPSTRI$, 6))
   On Error Resume Next
   MSComm1.CommPort = COPORT
   On Error GoTo 0
   '
   SETTI$ = ""
   For KKI% = 1 To 4
     SETTI$ = SETTI$ + TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
     If KKI% < 4 Then SETTI$ = SETTI$ + ","
   Next KKI%
   On Error Resume Next
   MSComm1.Settings = SETTI$
   If Err Then
      Call MENSERR(24, "Configuración de Puerto no Válida.\Corrija o Aplique Valores Predeterminados por Omisión.")
      Picture1.Visible = False
      SETUPTEL.Show 1
      Exit Sub
   End If
    '
    ' Abre el puerto de comunicaciones.
   On Error Resume Next
   MSComm1.PortOpen = True
   If Err Then
      Call MENSERR(24, "COM" + TRIM$(Str$(COPORT)) + " no está disponible.\Cambie la propiedad CommPort a otro puerto.")
      Picture1.Visible = False
      Exit Sub
   End If
    '
    ' Vuelca el búfer de entrada.
   MSComm1.InBufferCount = 0
    '
    ' Marca el número.
   MSComm1.Output = DialString$
   TFIN = Timer + 10
   Label30.Visible = False
    '
    ' Espera a que venga el "OK" del módem.
   Do
     dummy = DoEvents()
     CEVEN = MSComm1.CommEvent
     Select Case CEVEN
        Case comEvReceive
          Label29 = "Llamando ..."
          Label29.Refresh
          Label30.Visible = True
        Case comEvSend
          Label29 = "Marcando ..."
        Case Else
          'MsgBox "Error de Comunicacion"
          'exit do
     End Select
      '
      ' Si hay datos en el búfer, los lee.
     If MSComm1.InBufferCount Then
       FromModem$ = FromModem$ + MSComm1.Input
        ' Comprueba si hay "OK".
       If InStr(FromModem$, "OK") Then
             Label29 = "Llamando ..."
             Label29.Refresh
             Label30.Visible = True
             TFIN = Timer + 3
             While Timer < TFIN
               dummy = DoEvents()
             Wend
             Exit Do
          End If
       End If
        
       ' ¿El usuario ha elegido Cancelar?
       If cancelflag Then
          cancelflag = False
          Exit Do
       End If
       '
       If Timer > TFIN Then
         Picture1.Visible = False
         MsgBox "Imposible Establecer Comunicación"
         Exit Do
       End If
       '
    Loop
    '
    ' Desconecta el módem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Cierra el puerto.
    MSComm1.PortOpen = False
    Picture1.Visible = False
    Call FRESHALL
    '
    If INTERN$ <> "" Then
      TFIN = Timer + 3
      While Timer < TFIN: Wend
      Number$ = INTERN$
      INTERN$ = ""
      vuelta% = 2
      GoTo 10
    End If
    '
End Sub

Sub GRAFICON()
   '
5  IFICHA% = Val(Label26)
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Then REG& = 1 + ULTREG&("BSEGCOM")
   REGI% = REG&
   '
   If REG& <= ULTREG("BSEGCOM") Then
        RELE$ = REGLEIDO$("BSEGCOM", REG&)
      Else
        RELE$ = MKI$(REGI%) + Space$(1022)
   End If
   '
   Call REPLA(RELE$, Text4, 3, 62)      ' empresa
   Call REPLA(RELE$, Text5, 65, 64)     ' DIRECCION ELECTRONICA
   Call REPLA(RELE$, Text20, 129, 48)   ' CALLE Y NUMERO
   Call REPLA(RELE$, Text21, 177, 8)    ' C.POSTAL
   Call REPLA(RELE$, Text22, 185, 32)   ' LOCALIDAD
   Call REPLA(RELE$, Text19, 217, 32)   ' PROVINCIA
   Call REPLA(RELE$, Text23, 249, 24)   ' PAIS
   '
   Call REPLA(RELE$, Text18, 273, 24)   ' ORIGEN
   Call REPLA(RELE$, Text1, 297, 24)    ' RAMO
   Call REPLA(RELE$, Text2, 321, 48)    ' PRINC.PRODUCTO
   Call REPLA(RELE$, Text3, 369, 8) ' provincia particular
   Call REPLA(RELE$, Text26, 377, 24)   ' TELEFONO
   Call REPLA(RELE$, Text17, 401, 24) ' telef.partic.
   Call REPLA(RELE$, Text24, 425, 24) ' telef2.partic
   '
   For KKI% = 0 To 2
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, Text6(KKI%), 1, 32)    ' CONTACTO PERSONA 1 / 2 / 3
     Call REPLA(TTXX$, Text7(KKI%), 33, 24)
     Call REPLA(TTXX$, Text8(KKI%), 57, 24)
     Call REPLA(TTXX$, Text10(KKI%), 81, 16)
     Call REPLA(RELE$, TTXX$, 449 + 96 * KKI%, 96)
   Next KKI%
   '
   Call REPLA(RELE$, Text13, 737, 64) ' WEB-SITE
   
   PUNGRU$ = ""
   For U& = 1 To List2.ListCount
     PGR% = Val(Mid$(List2.List(U& - 1), 73, 6))
     If PGR% < 256 Then PUNGRU$ = PUNGRU$ + Chr$(PGR%)
   Next U&
   Call REPLA(RELE$, PUNGRU$ + String$(30, 0), 993, 30)
   '
   Call REPLA(RELE$, MKI$(FECHANUM(Label38)), 973, 2)   ' FECHA DE BAJA
   Call REPLA(RELE$, TRIM$(Text14), 975, 16)             ' MOTIVO DE BAJA
   Call REPLA(RELE$, MKI$(FECHANUM(Text12)), 1023, 2)   ' FECHA PROXIMO CONTACTO
   'HOII% = HOY(HOS$)
   'If CVI(Mid$(RELE$, 1023, 2)) < HOII% Then
   '  Call REPLA(RELE$, MKI$(HOII%), 1023, 2)
   'End If
   '
   Call GRAREG("BSEGCOM", RELE$, REG&)
   '
   Call CIERRARCH("BSEGCOM")
   Call FRESHECHO("BSEGCOM")
   '
   Screen.MousePointer = 11
   NCON% = Val(Label26)
   URN& = ULTREG&("NOTASEG")
   RNC$ = RECFILT$("NOTASEG", 1, MKI$(NCON%))
   KK& = (-3)
   '
   TTYZ$ = Text9.Text
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&)
   While Len(TEXTO$) > 0
      TTYY$ = MKI$(NCON%) + TEXTO$ + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("NOTASEG", TTYY$, RENOTA&)
      TEXTO$ = Mid$(TEXTO$, 63)
   Wend
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("NOTASEG", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("NOTASEG")
   Screen.MousePointer = 1

End Sub

Sub CARFICHA()
   '
5  IFICHA% = Val(Label26)
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BSEGCOM") Then
       Text1 = ""
       Text2 = ""
       Text3 = ""
       Text4 = ""
       Text5 = ""
       Text9 = ""
       Text16 = ""
       Text15 = ""
       Text13 = ""
       Text12 = ""
       Text25 = ""
       Text14 = ""
       Text11 = ""
       Text20 = ""
       Text21 = ""
       Text22 = ""
       Text19 = ""
       Text23 = ""
       Text26 = ""
       Text17 = ""
       Text24 = ""
       Text18 = ""
       '
       For KKI% = 0 To 2
         Text6(KKI%) = ""
         Text7(KKI%) = ""
         Text8(KKI%) = ""
         Text10(KKI%) = ""
       Next KKI%
       '
       HOII% = HOY(HOS$)
       Label34 = HOS$
       Text12 = FECHATEX$(DIPOST(HOII%))
       '
       List2.Clear
       '
     Else
       RELE$ = REGLEIDO$("BSEGCOM", REG&)
       Text4 = TRIM$(Mid$(RELE$, 3, 62))
       Text5 = TRIM$(Mid$(RELE$, 65, 64))
       Text20 = TRIM$(Mid$(RELE$, 129, 48))
       Text21 = TRIM$(Mid$(RELE$, 177, 8))
       '
       Text22 = TRIM$(Mid$(RELE$, 185, 32))
       Text19 = TRIM$(Mid$(RELE$, 217, 32))
       Text23 = TRIM$(Mid$(RELE$, 249, 24))
       Text18 = TRIM$(Mid$(RELE$, 273, 24))
       Text1 = TRIM$(Mid$(RELE$, 297, 24))
       Text2 = TRIM$(Mid$(RELE$, 321, 48))
       Text3 = TRIM$(Mid$(RELE$, 369, 8))
       '
       Text26 = TRIM$(Mid$(RELE$, 377, 24))
       Text17 = TRIM$(Mid$(RELE$, 401, 24))
       Text24 = TRIM$(Mid$(RELE$, 425, 24))
       '
       For KKI% = 0 To 2
         TTXX$ = Mid$(RELE$, 449 + KKI% * 96, 96)
         Text6(KKI%) = Mid$(TTXX$, 1, 32)
         Text7(KKI%) = Mid$(TTXX$, 33, 24)
         Text8(KKI%) = Mid$(TTXX$, 57, 24)
         Text10(KKI%) = Mid$(TTXX$, 81, 16)
       Next KKI%
       '
       Text13 = TRIM$(Mid$(RELE$, 737, 64))
       '
       List2.Clear
       PUNGRU$ = Mid$(RELE$, 993, 30)
       UTACON% = ULTREG&("TAGRUCON")
       For UKI% = 1 To Len(PUNGRU$)
         RETA& = Asc(Mid$(PUNGRU$, UKI%, 1))
         If RETA& > 0 Then
           If RETA& <= UTACON% Then
             XX$ = REGLEIDO$("TAGRUCON", RETA&)
             RAGRE% = Val(CVI(Left$(XX$, 2)))
             TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 24) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
             List2.AddItem TAGRE$
           End If
         End If
       Next UKI%
       '
       FEBAJA% = CVI(Mid$(RELE$, 973, 2))                ' FECHA DE BAJA
       Label38 = ""
       If FEBAJA% > 8224 Then Label38 = FECHATEX$(FEBAJA%)
       Text14 = Mid$(RELE$, 975, 16)                      ' MOTIVO DE BAJA
       '
       Text12 = FECHATEX$(CVI(Mid$(RELE$, 1023, 2)))
       '
       Text9 = "": TTYZ$ = ""
       If IFICHA% > 0 Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTASEG", 1, MKI$(IFICHA%))
         URN& = ULTREG&("NOTASEG")
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URN& Then
               TTYZ$ = TTYZ$ + Mid$(REGLEIDO$("NOTASEG", RENOTA&), 3)
             End If
           End If
         Next U&
         Screen.MousePointer = 1
       End If
       '
       LOCADE& = Len(TTYZ$)
       For U& = LOCADE& To 1 Step -1
         If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
       Next U&
       U& = 0
30     Text9 = Left$(TTYZ$, U&)
       Text9.Refresh
       '
   End If
   '
End Sub

Sub SORCONTAC()
   '
   OPXX% = COMALTER%("Elija si Criterio de Ordenamiento:\\Por Número\Alfabético\Cancelar")
   If OPXX% = 1 Then
       Call FILESORT("BSEGCOM", "", 1, 1, "ASC")
     ElseIf OPXX% = 2 Then
       Call FILESORT("BSEGCOM", "", 2, 2, "ASC")
   End If
   Call FRESHECHO("BSEGCOM")
   Exit Sub
   '
     
   List1.Clear
   Screen.MousePointer = 11
   '
   For U& = 1 To ULTREG&("BSEGCOM")
     RELE$ = REGLEIDO$("BSEGCOM", U&)
     TX2$ = UCase$(TRIM$(Mid$(RELE$, 3, 62)))
     TTXX$ = AJUSTI$(TX2$, 64) + Str$(U&)
     List1.AddItem TTXX$
   Next U&
   '
   JJ& = 0
   Call SETULTREG("!BSEGCOM", 0)
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     REG& = Val(Mid$(TTXX$, 65))
     X$ = REGLEIDO$("BSEGCOM", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!BSEGCOM", X$, JJ&)
   Next U&
   Call SETULTREG("!BSEGCOM", JJ&)
   Call CIERRARCH("*.*")
   '
   Call TRACENTRAL("BSEGCOM", "")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub EXPORAGE()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:EXPORLIB.CSV" For Output As #N1%
   Print #N1%, "Nombre,Dirección electrónica"
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     '
     TTYY$ = TRIM$(TTXX$ + TRIM$(TX4$))
     TTZZ$ = REPLACAR$(TTYY$, Chr$(34), "")
     TTXX$ = REPLACAR$(TTZZ$, ",", " ")
     '
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     EMA$ = TRIM$(Mid$(RELE$, 217, 64))
     If EMA$ <> "" Then
       If Asc(Left$(EMA$, 1)) <> 39 Then
         AX$ = TTXX$ + "," + EMA$
         Print #N1%, AX$
         JJ& = JJ& + 1
       End If
     End If
   Next U&
   '
   Close #N1%
   Call CIERRARCH("*.*")
   MsgBox Str$(JJ&) + " Direcciones Exportadas"
   Screen.MousePointer = 1
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Sub IMPORAGE()
   '
   Screen.MousePointer = 11
   Dim RAEMP$(4096)
   For U& = 1 To ULTREG&("BCONTACT")
     X$ = REGLEIDO$("BCONTACT", U&)
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
     CAEMP& = U&
   Next U&
   '
   FIN& = ULTREG&("!BCONTACT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!BCONTACT", U&)
     RASOE$ = UCase$(TRIM$(Mid$(X$, 145, 64)))
     EMAI$ = LCase$(TRIM$(Mid$(X$, 217, 64)))
     For JJ& = 1 To CAEMP&
       If RAEMP$(JJ&) = RASOE$ Then
         Y$ = REGLEIDO$("BCONTACT", JJ&)
         If InStr(Y$, EMAI$) > 0 Then GoTo 19 ' YA ESTA
         If TRIM$(Mid$(Y$, 217, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 217, 64)
           ElseIf TRIM$(Mid$(Y$, 281, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 281, 64)
           ElseIf TRIM$(Mid$(Y$, 345, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 345, 64)
         End If
         Call GRAREG("BCONTACT", Y$, JJ&)
         GoTo 19
       End If
    Next JJ&
    '
    CAEMP& = CAEMP& + 1
    RAEMP$(CAEMP&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
    KKI% = CAEMP&
    Call REPLA(X$, MKI$(KKI%), 1, 2)
    Call REGAPP("BCONTACT", X$)
19 Next U&
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub IMPOROUT()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:\OUTLOOK.CSV" For Input As #N1%
   Line Input #N1%, A$
   On Error GoTo 0
   '
   REG& = 0
   While Not EOF(N1%)
     REG& = REG& + 1
     REGI% = REG&
     RELE$ = MKI$(REGI%) + Space$(1022)
     '
     Line Input #N1%, TTXX$
     PPXX% = InStr(TTXX$, ",")
     If PPXX% > 0 Then
        EMPRE$ = Left$(TTXX$, PPXX% - 1)
        TTXX$ = Mid$(TTXX$, PPXX% + 1)
        PPXX% = InStr(TTXX$, ",")
        If PPXX% > 0 Then
          EMAI$ = Left$(TTXX$, PPXX% - 1)
          '
          Call REPLA(RELE$, EMPRE$, 3, 78)
          Call REPLA(RELE$, "Sr.", 81, 8)
          Call REPLA(RELE$, EMPRE$, 145, 64)
          Call REPLA(RELE$, EMAI$, 217, 64)
          Call GRAREG("BCONTACT", RELE$, REG&)
        End If
     End If
     '
   Wend
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Private Sub Text26_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text3_Change()
   NVENDE% = XVALO(Text3)
   Label40 = ""
   If NVENDE% < 256 Then
     Label40 = ECOARCH$("VENDEDOR", Chr$(NVENDE%))
   End If
End Sub

Private Sub Text3_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text4_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub Text5_DblClick()
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
     Call CREAMAIL(NPX&)
     '
     ATENTI% = 1
     MAILMESSVE.Label26 = Label26
     MAILMESSVE.Show 1
   End If
   '
End Sub

Private Sub Text5_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text7_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text8_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text9_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub ACTAGRUP()
   Call TRALOCAL("TAGRUCON", "")
   VTB% = VENTABU%("TAGRUCON")
   If VTB% >= 0 Then
     For U& = 1 To ULTREG&("!TAGRUCON")
       XX$ = REGLEIDO$("!TAGRUCON", U&)
       UKI% = U&
       Call REPLA(XX$, MKI$(UKI%), 1, 2)
       Call GRAREG("!TAGRUCON", XX$, U&)
     Next U&
     Call TRACENTRAL("TAGRUCON", "")
     If ULTREG&("TAGRUCON") > 255 Then
       Call SETULTREG("TAGRUCON", 255)
     End If
   End If
End Sub

Private Sub VScroll1_Change()
   

End Sub
