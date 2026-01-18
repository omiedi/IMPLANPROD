VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PLANCOTI09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Cotizaciones - Solicitud"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   177
      Top             =   7680
      Width           =   6975
      Begin VB.TextBox Text14 
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
         Left            =   240
         TabIndex        =   178
         Text            =   " "
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame10 
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
      Height          =   1335
      Left            =   1200
      TabIndex        =   170
      Top             =   7560
      Width           =   2655
      Begin VB.TextBox Text16 
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
         Left            =   1080
         TabIndex        =   173
         Text            =   " "
         Top             =   960
         Width           =   1455
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
         Left            =   1080
         TabIndex        =   172
         Text            =   " "
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text15 
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
         Left            =   1080
         TabIndex        =   171
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Relieve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   -360
         TabIndex        =   176
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Colores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   -360
         TabIndex        =   175
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stamping"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   -360
         TabIndex        =   174
         Top             =   660
         Width           =   1320
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Tintas"
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
      Left            =   4200
      TabIndex        =   162
      Top             =   7440
      Width           =   2175
      Begin VB.CheckBox Check10 
         Caption         =   "Check4"
         Height          =   255
         Left            =   1080
         TabIndex        =   168
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         Caption         =   "100%"
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
         Left            =   120
         TabIndex        =   167
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option12 
         Caption         =   "10%"
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
         Left            =   120
         TabIndex        =   166
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option11 
         Caption         =   "60%"
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
         Left            =   120
         TabIndex        =   165
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option8 
         Caption         =   "30%"
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
         Left            =   1080
         TabIndex        =   164
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "80%"
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
         Left            =   1080
         TabIndex        =   163
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Pantone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   1320
         TabIndex        =   169
         Top             =   960
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   7680
      TabIndex        =   157
      Top             =   7800
      Width           =   1455
      Begin VB.OptionButton Option30 
         Caption         =   "S/Datos"
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
         Index           =   2
         Left            =   120
         TabIndex        =   160
         Top             =   900
         Width           =   1095
      End
      Begin VB.OptionButton Option30 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option30 
         Caption         =   "Muestra"
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
         Index           =   1
         Left            =   120
         TabIndex        =   158
         Top             =   570
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Material"
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
      TabIndex        =   149
      Top             =   7560
      Width           =   7215
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
         Left            =   1200
         TabIndex        =   151
         Text            =   " "
         Top             =   360
         Width           =   5775
      End
      Begin VB.TextBox Text7 
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
         Left            =   1200
         TabIndex        =   150
         Text            =   " "
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adhesivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   -240
         TabIndex        =   153
         Top             =   435
         Width           =   1320
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Liner:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   -240
         TabIndex        =   152
         Top             =   900
         Width           =   1320
      End
   End
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
      Height          =   270
      Left            =   9480
      TabIndex        =   1
      Top             =   120
      Width           =   150
   End
   Begin VB.TextBox Text8 
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
      Left            =   1320
      TabIndex        =   2
      Text            =   " "
      Top             =   840
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   9900
      ScaleHeight     =   7755
      ScaleWidth      =   1155
      TabIndex        =   103
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Troq."
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
         Left            =   120
         Picture         =   "PLANCOTI09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Tabla de Troqueles"
         Top             =   3030
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Imprime Documento Seleccionado"
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4360
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
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
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3690
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Acceso a Base de Datos de Clientes"
         Top             =   2240
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         Picture         =   "PLANCOTI09.frx":1502
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Nueva Cotizacion"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Height          =   615
         Left            =   120
         Picture         =   "PLANCOTI09.frx":180C
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   6600
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":1B16
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   100
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Height          =   690
         Left            =   120
         Picture         =   "PLANCOTI09.frx":1C60
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   790
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   5880
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PLANCOTI09.frx":1F6A
         Top             =   6120
         Width           =   1515
      End
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
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10186
      _Version        =   327681
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos del Cliente"
      TabPicture(0)   =   "PLANCOTI09.frx":3E8C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame15"
      Tab(0).Control(1)=   "Frame19"
      Tab(0).Control(2)=   "Frame20"
      Tab(0).Control(3)=   "Frame24"
      Tab(0).Control(4)=   "Frame22"
      Tab(0).Control(5)=   "Frame21"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Producto"
      TabPicture(1)   =   "PLANCOTI09.frx":3EA8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label61"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label62"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text35"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame23"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame18"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame26"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame12"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text11"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Caracteristicas"
      TabPicture(2)   =   "PLANCOTI09.frx":3EC4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame14"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(7)=   "Frame25"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Cotizacion"
      TabPicture(3)   =   "PLANCOTI09.frx":3EE0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame29"
      Tab(3).Control(1)=   "Frame28"
      Tab(3).Control(2)=   "Frame27"
      Tab(3).Control(3)=   "Frame17"
      Tab(3).Control(4)=   "Frame16"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text11 
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
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "PLANCOTI09.frx":3EFC
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Frame Frame29 
         Caption         =   "Pre-Impresion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   185
         Top             =   2760
         Width           =   1935
         Begin VB.OptionButton Option33 
            Caption         =   "Incluir en Pre- cio por millar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   76
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Cotizar por Separado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   75
            Top             =   760
            Width           =   1335
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Datos de la Competencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -72720
         TabIndex        =   184
         Top             =   2760
         Width           =   7095
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
            Index           =   2
            Left            =   2160
            TabIndex        =   82
            Text            =   " "
            Top             =   1000
            Width           =   4695
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
            Index           =   0
            Left            =   2160
            TabIndex        =   78
            Text            =   " "
            Top             =   300
            Width           =   4695
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
            Index           =   1
            Left            =   2160
            TabIndex        =   80
            Text            =   " "
            Top             =   650
            Width           =   4695
         End
         Begin VB.TextBox Text23 
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
            Index           =   2
            Left            =   480
            TabIndex        =   81
            Text            =   " "
            Top             =   1000
            Width           =   1455
         End
         Begin VB.TextBox Text23 
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
            Index           =   0
            Left            =   480
            TabIndex        =   77
            Text            =   " "
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox Text23 
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
            Index           =   1
            Left            =   480
            TabIndex        =   79
            Text            =   " "
            Top             =   650
            Width           =   1455
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Resultado de Cotización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   182
         Top             =   4320
         Width           =   9135
         Begin VB.OptionButton Option9 
            Caption         =   "Pendiente"
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
            Left            =   240
            TabIndex        =   83
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox Text34 
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
            Left            =   2400
            MultiLine       =   -1  'True
            TabIndex        =   86
            Text            =   "PLANCOTI09.frx":3EFE
            Top             =   300
            Width           =   6615
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Perdido"
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
            Left            =   240
            TabIndex        =   85
            Top             =   800
            Width           =   1335
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Ganado"
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
            Left            =   240
            TabIndex        =   84
            Top             =   550
            Width           =   1335
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1080
            TabIndex        =   183
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "A Cotizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   179
         Top             =   520
         Width           =   2415
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   480
            TabIndex        =   71
            Text            =   " "
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   480
            TabIndex        =   63
            Text            =   " "
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   480
            TabIndex        =   59
            Text            =   " "
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   480
            TabIndex        =   67
            Text            =   " "
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   3840
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   165
            Left            =   480
            TabIndex        =   180
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   161
         Top             =   4920
         Width           =   3855
         Begin VB.OptionButton Option27 
            Caption         =   "Fotocromo"
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
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option28 
            Caption         =   "Pluma"
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
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option29 
            Caption         =   "Cod.Barras"
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
            Left            =   2400
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Incrustar Imagen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   5520
         TabIndex        =   154
         Top             =   600
         Width           =   3855
         Begin VB.CommandButton Command1 
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
            Left            =   3600
            TabIndex        =   31
            Top             =   3650
            Width           =   150
         End
         Begin VB.PictureBox Picture1 
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3075
            ScaleWidth      =   3555
            TabIndex        =   155
            Top             =   360
            Width           =   3615
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   2160
               Top             =   720
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Image Image1 
               Height          =   3135
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3615
            End
         End
         Begin VB.Label Label16 
            BackColor       =   &H00EEEEEE&
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
            Height          =   270
            Left            =   120
            TabIndex        =   156
            Top             =   3650
            Width           =   3495
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Condiciones a Soportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   148
         Top             =   3960
         Width           =   3975
         Begin VB.CheckBox Check21 
            Caption         =   "Otro"
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
            Left            =   3000
            TabIndex        =   28
            Top             =   250
            Width           =   855
         End
         Begin VB.CheckBox Check22 
            Caption         =   "Humedad"
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
            Left            =   1560
            TabIndex        =   27
            Top             =   250
            Width           =   1575
         End
         Begin VB.CheckBox Check20 
            Caption         =   "Frio"
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
            Left            =   480
            TabIndex        =   26
            Top             =   250
            Width           =   1335
         End
      End
      Begin VB.Frame Frame25 
         Height          =   1095
         Left            =   -67080
         TabIndex        =   145
         Top             =   1680
         Width           =   1575
         Begin VB.OptionButton Option32 
            Caption         =   "Dispensora"
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
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option32 
            Caption         =   "Impresora"
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
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option32 
            Caption         =   "Unidades"
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
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Orden"
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
         Left            =   1320
         TabIndex        =   144
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton Option17 
            Caption         =   "Repetición"
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
            Left            =   480
            TabIndex        =   19
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option16 
            Caption         =   "Nuevo"
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
            Left            =   480
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tecnica"
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
         Left            =   3480
         TabIndex        =   140
         Top             =   600
         Width           =   1815
         Begin VB.OptionButton Option2 
            Caption         =   "Digital"
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
            Left            =   360
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Flexografía"
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
            Left            =   360
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command12 
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
         Left            =   1800
         TabIndex        =   23
         Top             =   1560
         Width           =   150
      End
      Begin VB.TextBox Text35 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "PLANCOTI09.frx":3F00
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox Text4 
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
         Left            =   1320
         TabIndex        =   29
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3255
         TabIndex        =   30
         Top             =   4680
         Width           =   150
      End
      Begin VB.Frame Frame21 
         Caption         =   "Datos del Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   133
         Top             =   840
         Width           =   4455
         Begin VB.Label Label3 
            BackColor       =   &H00EEEEEE&
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
            Height          =   270
            Left            =   1080
            TabIndex        =   139
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -360
            TabIndex        =   138
            Top             =   540
            Width           =   1305
         End
         Begin VB.Label Label65 
            BackColor       =   &H00EEEEEE&
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
            Height          =   270
            Left            =   1080
            TabIndex        =   137
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -360
            TabIndex        =   136
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label Label67 
            BackColor       =   &H00EEEEEE&
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
            Height          =   270
            Left            =   1080
            TabIndex        =   135
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pais:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -360
            TabIndex        =   134
            Top             =   1500
            Width           =   1305
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Datos de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -70200
         TabIndex        =   129
         Top             =   840
         Width           =   4455
         Begin VB.TextBox Text37 
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
            Left            =   1200
            TabIndex        =   4
            Text            =   " "
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox Text38 
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
            Left            =   1200
            TabIndex        =   5
            Text            =   " "
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox Text39 
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
            Left            =   1200
            TabIndex        =   6
            Text            =   " "
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pais:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -240
            TabIndex        =   132
            Top             =   1500
            Width           =   1305
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -240
            TabIndex        =   131
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "L. Entrega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -600
            TabIndex        =   130
            Top             =   540
            Width           =   1665
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Datos Comerciales"
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
         Left            =   -74760
         TabIndex        =   125
         Top             =   2880
         Width           =   9015
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
            Height          =   290
            Left            =   1680
            TabIndex        =   8
            Top             =   375
            Width           =   150
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Left            =   1080
            TabIndex        =   7
            Text            =   " "
            Top             =   360
            Width           =   600
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
            Height          =   270
            Left            =   8760
            TabIndex        =   9
            Top             =   375
            Width           =   150
         End
         Begin VB.Label Label9 
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
            Height          =   270
            Left            =   1920
            TabIndex        =   35
            Top             =   375
            Width           =   3735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "    Vendedor: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   -240
            TabIndex        =   128
            Top             =   435
            Width           =   1320
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega Solicitada:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   5880
            TabIndex        =   127
            Top             =   435
            Width           =   1800
         End
         Begin VB.Label Label58 
            BackColor       =   &H00EEEEEE&
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
            Height          =   270
            Left            =   7800
            TabIndex        =   126
            Top             =   375
            Width           =   975
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Cond. Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -72720
         TabIndex        =   124
         Top             =   3840
         Width           =   2295
         Begin VB.OptionButton Option21 
            Caption         =   "Cta. Cte. Días"
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
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option22 
            Caption         =   "Anticipo y Contado"
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
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton Option23 
            Caption         =   "Contado C/Entrega"
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
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Cliente Privilegia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   123
         Top             =   3840
         Width           =   1935
         Begin VB.OptionButton Option18 
            Caption         =   "Entrega Rapida"
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
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option19 
            Caption         =   "Precio"
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
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option20 
            Caption         =   "Calidad"
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
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Antecedentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -70320
         TabIndex        =   122
         Top             =   3840
         Width           =   4575
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   16
            Text            =   "PLANCOTI09.frx":3F02
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Ultima Cotización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -72120
         TabIndex        =   119
         Top             =   520
         Width           =   6495
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
            Index           =   3
            Left            =   3960
            TabIndex        =   74
            Text            =   " "
            Top             =   1680
            Width           =   2415
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
            Index           =   2
            Left            =   3960
            TabIndex        =   70
            Text            =   " "
            Top             =   1320
            Width           =   2415
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
            Index           =   1
            Left            =   3960
            TabIndex        =   66
            Text            =   " "
            Top             =   960
            Width           =   2415
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
            Index           =   0
            Left            =   3960
            TabIndex        =   62
            Text            =   " "
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   240
            TabIndex        =   68
            Text            =   " "
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Text            =   " "
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   240
            TabIndex        =   64
            Text            =   " "
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   240
            TabIndex        =   72
            Text            =   " "
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   2160
            TabIndex        =   61
            Text            =   " "
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   2160
            TabIndex        =   65
            Text            =   " "
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   2160
            TabIndex        =   69
            Text            =   " "
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   2160
            TabIndex        =   73
            Text            =   " "
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   165
            Left            =   3960
            TabIndex        =   181
            Top             =   240
            Width           =   2400
         End
         Begin VB.Line Line1 
            X1              =   3840
            X2              =   3840
            Y1              =   120
            Y2              =   2040
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   165
            Left            =   240
            TabIndex        =   121
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Precio  X Millar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   165
            Left            =   2160
            TabIndex        =   120
            Top             =   240
            Width           =   1320
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   1920
            Y1              =   120
            Y2              =   2040
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   6480
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Medidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74640
         TabIndex        =   116
         Top             =   600
         Width           =   2535
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            TabIndex        =   37
            Text            =   " "
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   36
            Text            =   " "
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   1200
            TabIndex        =   118
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Avance"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   30
            TabIndex        =   117
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Troquel"
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
         Left            =   -74640
         TabIndex        =   114
         Top             =   1680
         Width           =   5775
         Begin VB.TextBox Text20 
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
            Left            =   960
            TabIndex        =   189
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command15 
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
            Left            =   2400
            TabIndex        =   188
            Top             =   360
            Width           =   150
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Nuevo"
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
            Left            =   2640
            TabIndex        =   45
            Top             =   360
            Width           =   975
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
            Left            =   120
            TabIndex        =   46
            Text            =   " "
            Top             =   720
            Width           =   5415
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stock:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   330
            Left            =   -120
            TabIndex        =   115
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Forma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72000
         TabIndex        =   113
         Top             =   600
         Width           =   3495
         Begin VB.OptionButton Option3 
            Caption         =   "Rectangular"
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
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Irregular"
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
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Circular / Ovalo"
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
            Left            =   1680
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Corte Medio"
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
            Left            =   1680
            TabIndex        =   41
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Forma de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -68400
         TabIndex        =   112
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton Option24 
            Caption         =   "Planchas"
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
            Left            =   1440
            TabIndex        =   43
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton Option26 
            Caption         =   "Indistinto"
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
            Left            =   960
            TabIndex        =   44
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option25 
            Caption         =   "Bobina"
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
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Aplicacion"
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
         Left            =   -68760
         TabIndex        =   111
         Top             =   1680
         Width           =   1695
         Begin VB.OptionButton Option40 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   320
            Width           =   975
         End
         Begin VB.OptionButton Option40 
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   48
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Terminacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74640
         TabIndex        =   110
         Top             =   2880
         Width           =   4455
         Begin VB.TextBox Text17 
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
            Height          =   765
            Left            =   1000
            MultiLine       =   -1  'True
            TabIndex        =   95
            Text            =   "PLANCOTI09.frx":3F04
            Top             =   1720
            Width           =   3255
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Otro:"
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
            Index           =   5
            Left            =   240
            TabIndex        =   57
            Top             =   1750
            Width           =   855
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Sectorizado"
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
            Index           =   3
            Left            =   240
            TabIndex        =   55
            Top             =   1190
            Width           =   1575
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   56
            Top             =   1470
            Width           =   1575
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Laminado"
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
            Index           =   2
            Left            =   240
            TabIndex        =   54
            Top             =   910
            Width           =   1335
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Barniz Agua"
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
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   350
            Width           =   1455
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Barniz UV"
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
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   630
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -70080
         TabIndex        =   109
         Top             =   2880
         Width           =   4575
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   160
            MultiLine       =   -1  'True
            TabIndex        =   58
            Text            =   "PLANCOTI09.frx":3F06
            Top             =   280
            Width           =   4215
         End
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cambios:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   -120
         TabIndex        =   186
         Top             =   3240
         Width           =   1320
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   -120
         TabIndex        =   147
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label Label17 
         BackColor       =   &H00EEEEEE&
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
         Height          =   495
         Left            =   1320
         TabIndex        =   146
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label Label62 
         BackColor       =   &H00EEEEEE&
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
         Height          =   270
         Left            =   2040
         TabIndex        =   143
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   -120
         TabIndex        =   142
         Top             =   1635
         Width           =   1320
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   450
         Left            =   -120
         TabIndex        =   141
         Top             =   4560
         Width           =   1320
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6120
      Picture         =   "PLANCOTI09.frx":3F08
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   6120
      Picture         =   "PLANCOTI09.frx":4212
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label75 
      BackColor       =   &H00EEEEEE&
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
      Height          =   270
      Left            =   1320
      TabIndex        =   108
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label74 
      BackColor       =   &H00EEEEEE&
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
      Height          =   270
      Left            =   4440
      TabIndex        =   107
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Prospecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   210
      Left            =   3360
      TabIndex        =   106
      Top             =   180
      Width           =   960
   End
   Begin VB.Label Label12 
      BackColor       =   &H00EEEEEE&
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
      Height          =   270
      Left            =   8040
      TabIndex        =   102
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Cotización:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6240
      TabIndex        =   101
      Top             =   180
      Width           =   1665
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reventa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   330
      Left            =   6240
      TabIndex        =   100
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -120
      TabIndex        =   99
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00004040&
      Height          =   330
      Left            =   7560
      TabIndex        =   98
      Top             =   540
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00EEEEEE&
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
      Height          =   270
      Left            =   8640
      TabIndex        =   97
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EEEEEE&
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
      Height          =   270
      Left            =   1320
      TabIndex        =   96
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00004040&
      Height          =   210
      Left            =   510
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "PLANCOTI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIMAGEN$
'
Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   '
   If TRIM$(Label12) = "" Then GoTo 99
   '
   CommonDialog1.filename = ""
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Imagenes"
   CommonDialog1.FILTER = "Imagen Tipo(*.JPG;*.BMP;*.GIF)|*.JPG|*.BMP|*.GIF|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   Label16 = UCase$(CommonDialog1.FileTitle)
   DIRECC$ = CommonDialog1.filename
   Image1.Picture = LoadPicture(DIRECC$)
   TIMAGEN$ = LOADFILE$(DIRECC$)
        
    Picture1.Picture = LoadPicture("")
    Picture1.Refresh
    '
    HIMAGEN = Image1.Picture.Width
    VIMAGEN = Image1.Picture.Height

    HCUADRO = Picture1.Width
    VCUADRO = Picture1.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
    Picture1.Height = VCUADRO * REDUV / REDUC
    Picture1.Width = HCUADRO * REDUH / REDUC
    Picture1.Refresh
    '
    HCUADRO = Picture1.Width
    VCUADRO = Picture1.Height
    Picture1.Left = (Frame26.Width - HCUADRO) / 2
    Picture1.Top = (Frame26.Height - VCUADRO) / 2
    '
    Picture1.PaintPicture Image1.Picture, 0, 0, HCUADRO, VCUADRO
    Picture1.Visible = True
    Picture1.Refresh
   
   'Call GRABAFOTO
   '
99 Command1.Enabled = True
   '
End Sub

Private Sub Command10_Click()
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
       Call CONECRUN("SEGCOM07", "Seguimiento Comercial")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call SELECHO("GRUCOVEN")
   VALONUE$ = VALACT1$("GRUCOVEN")
   If VALONUE$ <> "" Then
      Text35.TEXT = VALONUE$
   End If
End Sub

Private Sub Command13_Click()
   '
   FORIPRE$ = TRIM$(CONTROL$("", "FORCOTI"))
   If FORIPRE$ = "" Then Exit Sub
   '
   If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito por Rechazo:\  \'" + RMCC$ + "'.")
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   '
   FEX = HOY(HOS$)
   FECHDOC$ = FECHATEX$(FEX)
   NUMEDOC$ = TRIM$(Label12)
   TIPODOC$ = "Solicitud de Cotizacion"
   '
   
   CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = TRIM$(Label12)
   CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHDOC$
   '
   '
   CODPARAM$(3) = "RASO-CLI": DOCPARAM$(3) = TRIM$(Label2)
   CODPARAM$(4) = "DOMI-CLI": DOCPARAM$(4) = TRIM$(Label3)
   CODPARAM$(5) = "LOCA-CLI": DOCPARAM$(5) = TRIM$(Label65)
   CODPARAM$(6) = "PAIS-CLI": DOCPARAM$(6) = TRIM$(Label67)
   CODPARAM$(7) = "CTAC-CLI": DOCPARAM$(7) = TRIM$(Text8)
   CODPARAM$(8) = "NRO-PASE": DOCPARAM$(8) = TRIM$(Text37)
   CODPARAM$(9) = "REF-MAT3": DOCPARAM$(9) = TRIM$(Text38)
   CODPARAM$(10) = "REF-MAT4": DOCPARAM$(10) = TRIM$(Text39)
   CODPARAM$(11) = "N-PLANO": DOCPARAM$(11) = TRIM$(Text3)
   CODPARAM$(12) = "VENDEDOR": DOCPARAM$(12) = TRIM$(Label9)
   CODPARAM$(13) = "FECH-VEN": DOCPARAM$(13) = TRIM$(Label58)
   CODPARAM$(14) = "REF-MAT2": DOCPARAM$(14) = Mid$(VALEIDCOT(TRIM$(Label12), "CLIENTEPRIVILEGIA"), 5)
   CODPARAM$(15) = "COND-PAG": DOCPARAM$(15) = Mid$(VALEIDCOT(TRIM$(Label12), "CONDPAGO"), 5)
   '
   '
   CODPARAM$(16) = "R-PLANO": DOCPARAM$(16) = Text19
   CODPARAM$(17) = "TIPO-VAL": DOCPARAM$(17) = Mid$(VALEIDCOT(TRIM$(Label12), "REPITE"), 5)
   CODPARAM$(18) = "REF-ASIE": DOCPARAM$(18) = Mid$(VALEIDCOT(TRIM$(Label12), "TIPO"), 5)
   CODPARAM$(19) = "REF-MAT1": DOCPARAM$(19) = TRIM$(Label62)
   CODPARAM$(20) = "BANC-CHE": DOCPARAM$(20) = Text11
   CODPARAM$(21) = "NUME-CHE": DOCPARAM$(21) = Text5
   CODPARAM$(22) = "DES-MAT": DOCPARAM$(22) = TRIM$(Label17)
   CODPARAM$(23) = "COD-MAT": DOCPARAM$(23) = TRIM$(Text4)
   '
   If TRIM$(VALEIDCOT(TRIM$(Label12), "SOPORTAFRIO")) = "SI" Then SOPOFRI$ = "Frio,"
   If TRIM$(VALEIDCOT(TRIM$(Label12), "SOPORTAHUMEDAD")) = "SI" Then SOPOHUM$ = "Humedad, "
   If TRIM$(VALEIDCOT(TRIM$(Label12), "SOPORTAOTRO")) = "SI" Then SOPOTRO$ = "Otro."
   SOPO$ = TRIM$(SOPOFRI$) & " " & TRIM$(SOPOHUM$) & " " & TRIM$(SOPOTRO$)
   CODPARAM$(24) = "MEDIDA-3": DOCPARAM$(24) = SOPO$
   '
   CODPARAM$(25) = "DES-CONJ": DOCPARAM$(25) = Mid$(VALEIDCOT(TRIM$(Label12), "TIPOIMPRESION"), 5)
   '
   CODPARAM$(26) = "MEDIDA-1": DOCPARAM$(26) = TRIM$(Text10)
   CODPARAM$(27) = "MEDIDA-2": DOCPARAM$(27) = TRIM$(Text9)
   CODPARAM$(28) = "DESTINO": DOCPARAM$(28) = Mid$(VALEIDCOT(TRIM$(Label12), "FORMA"), 4)
   CODPARAM$(29) = "REF-MAT5": DOCPARAM$(29) = Mid$(VALEIDCOT(TRIM$(Label12), "FORMAENTREGA"), 5)
   '
   APLICA1$ = Mid$(VALEIDCOT(TRIM$(Label12), "APLICACION_1"), 4)
   If APLICA1$ <> "" Then
      APLICA2$ = ", " & Mid$(VALEIDCOT(TRIM$(Label12), "APLICACION_2"), 4)
   Else
      APLICA2$ = Mid$(VALEIDCOT(TRIM$(Label12), "APLICACION_2"), 4)
   End If
   CODPARAM$(30) = "REF-MAT6": DOCPARAM$(30) = TRIM$(APLICA1$) & TRIM$(APLICA2$)
   '
   If Option31(5).Value = True Then
        CODPARAM$(31) = "TIP-MPRI": DOCPARAM$(31) = "Otro: " & Mid$(VALEIDCOT(TRIM$(Label12), "TERMINACION"), 5)
   Else
        CODPARAM$(31) = "TIP-MPRI": DOCPARAM$(31) = Mid$(VALEIDCOT(TRIM$(Label12), "TERMINACION"), 5)
   End If
'   If VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO") = "NUEVO" Then
'        TERMINA$ = "NUEVO" & " - " & Text12
'   ElseIf VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO") = "" Then
'        TERMINA$ = Text12
'   End If
  If VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO") = "NUEVO" Then
        TERMINA$ = "NUEVO" & " - " & Text20
   ElseIf VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO") = "" Then
        TERMINA$ = Text20
   End If
   
   
   CODPARAM$(32) = "DOCU-APL": DOCPARAM$(32) = TERMINA$
   '
   CODPARAM$(33) = "OBSERVA": DOCPARAM$(33) = Text18
   '
   CODPARAM$(34) = "TIT-MPRI": DOCPARAM$(34) = Mid$(VALEIDCOT(TRIM$(Label12), "PREIMPRESION"), 5)
   CODPARAM$(35) = "COL-MPRI": DOCPARAM$(35) = Mid$(VALEIDCOT(TRIM$(Label12), "RESULTADO"), 5)
   CODPARAM$(36) = "PRO-MPRI": DOCPARAM$(36) = Text34
  
   
   CAPARDOC% = 36
   '
   

   CODTABU1$(1) = "UNI-MPRI"
   CODTABU1$(2) = "CAN-MPRI"
   CODTABU1$(3) = "USO-MPRI"
   CODTABU1$(4) = "DES-MPRI"
   CODTABU1$(5) = "UBI-MPRI"
   CODTABU1$(6) = "PAR-MPRI"
   CACOLTAB1% = 6
   '
   CAITAB1% = 0
   For YX& = 0 To 3
   
      CANTCO$ = TRIM$(Text22(YX&))
      CANULT$ = TRIM$(Text21(YX&))
      PRMIL$ = TRIM$(Text24(YX&))
      OBSERR$ = TRIM$(Text1(YX&))
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = CANTCO$
      TETABU1$(2, CAITAB1%) = CANULT$
      TETABU1$(3, CAITAB1%) = PRMIL$
      TETABU1$(4, CAITAB1%) = OBSERR$
      '
   Next YX&
   For YX& = 0 To 2
      DATOCOMP1$ = TRIM$(Text25(YX&))
      DATOCOMP2$ = TRIM$(Text23(YX&))
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(5, CAITAB1%) = DATOCOMP1$
      TETABU1$(6, CAITAB1%) = DATOCOMP2$
      '
   Next YX&

   '
   Call PRINTDOC("@SOCO-ACN")   ' SOLICITUD DE COTIZACION
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command14_Click()
   '
   Call SELECHO("SELFECHA")
   FECCH$ = VALACT1$("SELFECHA")
   If TRIM$(FECCH$) <> "" Then
        Label58 = FECCH$
   End If
   '
End Sub

Private Sub Command15_Click()
    '
    Call SELECHO("INDITROQ")
       TROQ$ = VALACT1$("INDITROQ")
    If TRIM$(TROQ$) <> "" Then
        Text20.TEXT = TRIM$(VALACT1$("INDITROQ"))
        Text12.TEXT = TRIM$(VALACT2$("INDITROQ"))
    End If
    '
End Sub

Private Sub Command3_Click()
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   CCII% = (CODINT%(COD$))
   If CCII% > 0 Then
     Text4.TEXT = CODEXT$(CCII%)
     Label17 = DESCRIT$(CCII%)
   End If
   '
End Sub

Private Sub Command2_Click()
    '
10  OPX% = ALTERNA%("Nuevo Troquel\Modificar Troquel")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 1 Then Call CARTROQUE
    If OPX% = 2 Then Call MODTROQUE
    GoTo 10
    '
End Sub
 Sub CARTROQUE()
    '
    VDEF$ = Space$(256): VDEF$ = ""
10  OBX$ = OBJPLA$("CARGATROQUEL", VDEF$)
    '
    If OBX$ = "" Then Exit Sub
    '
    CODTROQ$ = TRIM$(Mid$(OBX$, 1, 12))
    DENOTROQ$ = TRIM$(Mid$(OBX$, 13, 64))
    FORMTROQ$ = TRIM$(Mid$(OBX$, 77, 32))
    AVANCC = CVS(Mid$(OBX$, 109, 4))
    ANCHH = CVS(Mid$(OBX$, 113, 4))
    DIAMM = CVS(Mid$(OBX$, 117, 4))
    BOCAS% = CVI(Mid$(OBX$, 121, 2))
    MEDIDD = CVS(Mid$(OBX$, 123, 4))
    AN = CVS(Mid$(OBX$, 127, 4))
    C10$ = TRIM$(Mid$(OBX$, 131, 12))
    C7$ = TRIM$(Mid$(OBX$, 143, 12))
    ENG = CVS(Mid$(OBX, 155, 4))
    SUPERFI = CVS(Mid$(OBX$, 159, 4))
    '
    Call REPLA(VDEF$, CODTROQ, 1, 12)
    Call REPLA(VDEF$, DENOTROQ$, 13, 64)
    Call REPLA(VDEF$, FORMTROQ$, 77, 32)
    Call REPLA(VDEF$, MKS$(AVANCC), 109, 4)
    Call REPLA(VDEF$, MKS$(ANCHH), 113, 4)
    Call REPLA(VDEF$, MKS$(DIAMM), 117, 4)
    Call REPLA(VDEF$, MKI$(BOCAS%), 121, 2)
    Call REPLA(VDEF$, MKS$(MEDIDD), 123, 4)
    Call REPLA(VDEF$, MKS$(AN), 127, 4)
    Call REPLA(VDEF$, C10$, 131, 12)
    Call REPLA(VDEF$, C7$, 143, 12)
    Call REPLA(VDEF$, MKS$(ENG), 155, 4)
    Call REPLA(VDEF$, MKS$(SUPERFI), 159, 4)
    '
    If TRIM$(CODTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Codigo del Troquel")
        GoTo 10
    End If
    '
    If TRIM$(DENOTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Descripcion del Troquel")
        GoTo 10
    End If
    '
    If TRIM$(FORMTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Forma del Troquel")
        GoTo 10
    End If
    
    If XVALO(AVANCC) = 0 Then
        Call MENSERR(24, "Debe Ingresar Avance del Troquel")
        GoTo 10
    End If
    If XVALO(ANCHH) = 0 Then
        Call MENSERR(24, "Debe Ingresar Ancho del Troquel")
        GoTo 10
    End If
    '
    Call ABREPEDCLI
    sqlx$ = "SELECT * FROM TROQUELES"
    
    Dim TROQUELTMP As ADODB.Recordset
    Set TROQUELTMP = New ADODB.Recordset
    TROQUELTMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    With TROQUELTMP
        .AddNew   ' YA CONVERTIDO
        ON ERROR RESUME NEXT
        !CodiEmpr = Codiemp%
        ON ERROR GOTO 0
            !CodTroquel = CODTROQ$
            DESTROQUEL$ = FORMATNUM$(AVANCC, "F5.1") & " X " & FORMATNUM$(ANCHH, "F5.1") & " - " & TRIM$(FORMTROQ$) & " - " & TRIM$(DENOTROQ$)
            !DescriTroquel = TRIM$(DESTROQUEL$)
            !DenoTroquel = DENOTROQ$
            !FormTroquel = FORMTROQ$
            !AvanceTroquel = AVANCC
            !AnchoTroquel = ANCHH
            !DiamTroquel = DIAMM
            !BocasTroquel = BOCAS%
            !MediTroquel = MEDIDD
        .Update
    End With
    Call MENSERR(24, "Troquel Cargado")
    
 End Sub
Sub MODTROQUE()
    '
    Call COPYSTRU("INDICOTI", "INDITROQ")
    Call CARGALISEL("INDITROQ", "TROQUELES", "CODTROQUEL/A17;DESCRITROQUEL/A48", "CODTROQUEL <> ''")
    Call SELECHO("INDITROQ")
       TROQ$ = VALACT1$("INDITROQ")
   If TRIM$(TROQ$) <> "" Then
    CONDI$ = "CODTROQUEL = '" & TRIM$(TROQ$) & "'"
    '
    CODTROQ$ = VALOBADA("TROQUELES", "CODTROQUEL", CONDI$, "NOMESS")
    DENOTROQ$ = VALOBADA("TROQUELES", "DenoTroquel", CONDI$, "NOMESS")
    FORMTROQ$ = VALOBADA("TROQUELES", "FormTroquel", CONDI$, "NOMESS")
    AVANCC = VALOBADA("TROQUELES", "AvanceTroquel", CONDI$, "NOMESS")
    ANCHH = VALOBADA("TROQUELES", "AnchoTroquel", CONDI$, "NOMESS")
    DIAMM = VALOBADA("TROQUELES", "DiamTroquel", CONDI$, "NOMESS")
    BOCAS% = XVALO(VALOBADA("TROQUELES", "BocasTroquel", CONDI$, "NOMESS"))
    MEDIDD = VALOBADA("TROQUELES", "MediTroquel", CONDI$, "NOMESS")
    '
    Call REPLA(VDEF$, CODTROQ, 1, 12)
    Call REPLA(VDEF$, DENOTROQ$, 13, 64)
    Call REPLA(VDEF$, FORMTROQ$, 77, 32)
    Call REPLA(VDEF$, MKS$(AVANCC), 109, 4)
    Call REPLA(VDEF$, MKS$(ANCHH), 113, 4)
    Call REPLA(VDEF$, MKS$(DIAMM), 117, 4)
    Call REPLA(VDEF$, MKI$(BOCAS%), 121, 2)
    Call REPLA(VDEF$, MKS$(MEDIDD), 123, 4)
    Call REPLA(VDEF$, MKS$(AN), 127, 4)
    Call REPLA(VDEF$, C10$, 131, 12)
    Call REPLA(VDEF$, C7$, 143, 12)
    Call REPLA(VDEF$, MKS$(ENG), 155, 4)
    Call REPLA(VDEF$, MKS$(SUPERFI), 159, 4)
    '
10  OBX$ = OBJPLA$("CARGATROQUEL", VDEF$)
    '
    If OBX$ = "" Then Exit Sub
    '
    CODTROQ$ = TRIM$(Mid$(OBX$, 1, 12))
    DENOTROQ$ = TRIM$(Mid$(OBX$, 13, 64))
    FORMTROQ$ = TRIM$(Mid$(OBX$, 77, 32))
    AVANCC = CVS(Mid$(OBX$, 109, 4))
    ANCHH = CVS(Mid$(OBX$, 113, 4))
    DIAMM = CVS(Mid$(OBX$, 117, 4))
    BOCAS% = CVI(Mid$(OBX$, 121, 2))
    MEDIDD = CVS(Mid$(OBX$, 123, 4))
    AN = CVS(Mid$(OBX$, 127, 4))
    C10$ = TRIM$(Mid$(OBX$, 131, 12))
    C7$ = TRIM$(Mid$(OBX$, 143, 12))
    ENG = CVS(Mid$(OBX, 155, 4))
    SUPERFI = CVS(Mid$(OBX$, 159, 4))
    '
    If TRIM$(CODTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Codigo del Troquel")
        GoTo 10
    End If
    '
    If TRIM$(DENOTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Descripcion del Troquel")
        GoTo 10
    End If
    '
    If TRIM$(FORMTROQ$) = "" Then
        Call MENSERR(24, "Debe Ingresar Forma del Troquel")
        GoTo 10
    End If
    
    If XVALO(AVANCC) = 0 Then
        Call MENSERR(24, "Debe Ingresar Avance del Troquel")
        GoTo 10
    End If
    If XVALO(ANCHH) = 0 Then
        Call MENSERR(24, "Debe Ingresar Ancho del Troquel")
        GoTo 10
    End If
    
    '
    Call ACTUREGISTRO("troqueles", "CODTROQUEL", CONDI$, CODTROQ$, REG&, "NOMESS")
    DESTROQUEL$ = FORMATNUM$(AVANCC, "F5.1") & " X " & FORMATNUM$(ANCHH, "F5.1") & " - " & TRIM$(FORMTROQ$) & " - " & TRIM$(DENOTROQ$)
    Call ACTUREGISTRO("troqueles", "DescriTroquel", CONDI$, TRIM$(DESTROQUEL$), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "DeNOTroquel", CONDI$, TRIM$(DENOTROQ$), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "FormTroquel", CONDI$, TRIM$(FORMTROQ$), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "AvanceTroquel", CONDI$, FORMATNUM$(AVANCC, "F5.1"), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "AnchoTroquel", CONDI$, FORMATNUM$(ANCHH, "F5.1"), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "DiamTroquel", CONDI$, FORMATNUM$(DIAMM, "F5.1"), REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "BocasTroquel", CONDI$, BOCAS%, REG&, "NOMESS")
    Call ACTUREGISTRO("troqueles", "MediTroquel", CONDI$, FORMATNUM$(MEDIDD, "F5.1"), REG&, "NOMESS")
    '
    Call MENSERR(24, "Troquel Modificado")
    '
   End If
    '

End Sub
    


Sub GRACOTIZA()

'    Call ABREPEDCLI
'    sqlx$ = "SELECT * FROM COTIZA"
'
'    Dim COTIZATMP As ADODB.Recordset
'    Set COTIZATMP = New ADODB.Recordset
'    COTIZATMP.Open sqlx$, FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'    With COTIZATMP
'        .AddNew
'            !Nro_Coti = XVALO(Label12)
'            !FECHA = Label4
'            !Codi_Cli = XVALO(Text1)
'            !Vendedor = XVALO(Text2)
'            !Contacto = TRIM$(Text8)
'            If Option1.Value = True Then
'                !Tipo = 1  'FLEXO
'              Else
'                !Tipo = 2  'DIGITAL
'            End If
'            If Option17.Value = True Then
'                !Nuevo = 0  'REPETICION
'              Else
'                !Nuevo = 1  'NUEVO
'            End If
'            '
'            !Cod_inte = CODINT(XVALO(Text4))
'            !Descripcion = TRIM$(Text5)
'            !Adhesivo = TRIM$(Text6)
'            !Liner = TRIM$(Text7)
'            '
'            If Check2.Value = 1 Then
'                !Referencia = 1     'Print
'            ElseIf Check1.Value = 1 Then
'                !Referencia = 2     'Muestra
'            ElseIf Check3.Value = 1 Then
'                !Referencia = 3     'S/Datos
'            End If
'            '
'            !Mid_Avance = XVALO(Text10)
'            !Mid_ancho = XVALO(Text9)
'            !TroquelNuevo = XVALO(Text11)
'            !TroquelStock = XVALO(Text12)
'            '
'            If Option3.Value = True Then
'                !FORMA = 1  'RECTANGULAR
'            ElseIf Option5.Value = True Then
'                !FORMA = 2  'Circular / Ovalo
'            ElseIf Option4.Value = True Then
'                !FORMA = 3  'Irregular
'            ElseIf Option6.Value = True Then
'                !FORMA = 4  'Corte Medio
'            End If
'            '
'            If Check22.Value = 1 Then
'                !FormaEntrega = 1     'Bobinas
'            ElseIf Check23.Value = 1 Then
'                !FormaEntrega = 2     'Planchas
'            ElseIf Check24.Value = 1 Then
'                !FormaEntrega = 3     'Indistinto
'            End If
'            '
'            If Option17.Value = True Then
'                !Aplicacion_1 = 1  'Manual
'              Else
'                !Aplicacion_1 = 2  'Automatica
'            End If
'            '
'            If Check7.Value = 1 Then
'                !Aplicacion_2 = 1     'Impresora
'            ElseIf Check9.Value = 1 Then
'                !Aplicacion_2 = 2     'Dispensora
'            ElseIf Check8.Value = 1 Then
'                !Aplicacion_2 = 3     'Unidades
'            End If
'            '
'            !Impre_Colores = TRIM$(Text15)
'            !Impre_Stamping = TRIM$(Text13)
'            !Impre_Relieve = TRIM$(Text16)
'
'            If Check10.Value = 1 Then
'                !Pantone = 1     'Si
'            Else
'                !Pantone = 0     'No
'            End If
'            '
'            If Option12.Value = True Then
'                !Tinta = 1  '10%
'              ElseIf Option8.Value = True Then
'                !Tinta = 2  '30%
'              ElseIf Option11.Value = True Then
'                !Tinta = 3  '60%
'              ElseIf Option7.Value = True Then
'                !Tinta = 4  '80%
'              ElseIf Option13.Value = True Then
'                !Tinta = 5  '100%
'            End If
'            '
'            If Check16.Value = 1 Then
'                !Terminacion = TRIM$(Label37)     'Barniz Agua
'            ElseIf Check15.Value = 1 Then
'                !Terminacion = TRIM$(Label35)     'Laminado
'            ElseIf Check18.Value = 1 Then
'                !Terminacion = TRIM$(Label39)     'Ninguno
'            ElseIf Check14.Value = 1 Then
'                !Terminacion = TRIM$(Label36)     'Barniz Uv
'            ElseIf Check17.Value = 1 Then
'                !Terminacion = TRIM$(Label38)     'Sectorizado
'            ElseIf Check19.Value = 1 Then
'                !Terminacion = TRIM$(Text17)      'Otros
'            End If
'            '
'            !Observaciones = TRIM$(Text18)
'            '
'            If Check20.Value = 1 Then
'                !Condiciones = 1     'Frio
'            ElseIf Check22.Value = 1 Then
'                !Condiciones = 2     'Humedad
'            ElseIf Check21.Value = 1 Then
'                !Condiciones = 3     'otros
'            End If
'            '
'            If Check23.Value = 1 Then
'                !Privilegia = 1     'Precio
'            ElseIf Check25.Value = 1 Then
'                !Privilegia = 2     'Calidad
'            ElseIf Check24.Value = 1 Then
'                !Privilegia = 3     'Entrega
'            End If
'            '
'            If Check26.Value = 1 Then
'                !Privilegia = 1     'anticipo
'            ElseIf Check28.Value = 1 Then
'                !Privilegia = 2     'Contado C/entrega
'            ElseIf Check27.Value = 1 Then
'                !Privilegia = 3     'Cta.Cte.Dias
'            End If
'            '
'            !Antecedentes = TRIM$(Text19)
'            !Cantidad_1 = XVALO(Text21)
'            !Precio_1 = XVALO(Text24)
'            !Cantidad_2 = XVALO(Text22)
'            !Precio_2 = XVALO(Text25)
'            !Cantidad_3 = XVALO(Text20)
'            !Precio_3 = XVALO(Text26)
'            !Cantidad_4 = XVALO(Text23)
'            !Precio_4 = XVALO(Text27)
'            !Arm_Foto = TRIM$(Text31)
'            !Troquel = TRIM$(Text30)
'            !Stamping = TRIM$(Text29)
'            !Relieve = TRIM$(Text28)
'            !PlazoEntrega = TRIM$(Text32)
'        .Update
'     End With

End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    If DERACCE%("MODICOTI", "Modificar Cotizaciones") < 2 Then
        GoTo 99
    End If
    '
    If Label12.Caption = "" Then GoTo 99
    '
    Call ABREPEDCLI
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "'"
    '
    ESTTD% = XVALO(VALOBADA("PRESUVEN", "Status", CONDI$, "NOMESS"))
    If ESTTD > 0 Then
        Call MENSERR(24, "La Cotización se Encuentra Cerrada\ \Debe Abrirla para Guardar las Modificaciones")
        GoTo 99
    End If
    '
    If TRIM$(Text20) <> "" And CantRegistros("TROQUELES", "CodTroquel = '" & TRIM(Text20) & "'") < 1 Then
        Call MENSERR(24, "Troquel no Existente")
        GoTo 99
    End If
    '
    Call ACTUREGISTRO("PRESUVEN", "Vended", CONDI$, XVALO(Text2), REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "AteSen", CONDI$, TRIM$(Text8), REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "Reventa", CONDI$, TRIM$(Text3), REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "CodiInt", CONDI$, CODINT%(Text4), REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "Especif", CONDI$, Text5, REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "Antecedentes", CONDI$, Text19, REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "Observa", CONDI$, Text18, REG&, "NOMESS")
    Call ACTUREGISTRO("PRESUVEN", "Imagen", CONDI$, TIMAGEN$, REG&, "NOMESS")
    '
    'DATOS DEL CLIENTE
    '
    Call GUARDAVALCOT(TRIM$(Label12), "LUGARENTREGA", Text37)
    Call GUARDAVALCOT(TRIM$(Label12), "LOCAENTREGA", Text38)
    Call GUARDAVALCOT(TRIM$(Label12), "PAISENTREGA", Text39)
    'Call GUARDAVALCOT(TRIM$(Label12), "VENDEDOR", Text2) LO TOMA DE PRESUVEN
    '
    'GUARDA SEGUN PRIVILEGIA EL CLIENTE
    If Option19.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CLIENTEPRIVILEGIA", "1 - PRECIO")   'PRECIO
    If Option20.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CLIENTEPRIVILEGIA", "2 - CALIDAD")   'CALIDAD
    If Option18.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CLIENTEPRIVILEGIA", "3 - ENTREGA RAPIDA")   'ENTREGA RAPIDA
    '
    'GUARDA LA CONDICIONES DE PAGO
    If Option22.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CONDPAGO", "1 - ANTICIPO Y CONTADO")   'ANTICIPO Y CONTADO
    If Option23.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CONDPAGO", "2 - CONTADO C/ENTREGA")   'CONTADO C/ENTREGA
    If Option21.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "CONDPAGO", "3 - Cta. Cte. Días")   'Cta. Cte. Días
    '
    Call GUARDAVALCOT(TRIM$(Label12), "FECHAENTREGA", Label58)
    '
    'CARACTERISTICA DEL PRODUCTO
    '
    Call GUARDAVALCOT(TRIM$(Label12), "MEDIDAAVANCE", Text10)
    Call GUARDAVALCOT(TRIM$(Label12), "MEDIDAANCHO", Text9)
    
    If Check1.Value = 1 Then
        Call GUARDAVALCOT(TRIM$(Label12), "TROQUELNUEVO", "NUEVO")
        Else
        Call GUARDAVALCOT(TRIM$(Label12), "TROQUELNUEVO", "")
    End If
'    If TRIM$(Text12) <> "" Then
'        Call GUARDAVALCOT(TRIM$(Label12), "TROQUELSTOCK", Text12)
'    End If
    If TRIM$(Text20) <> "" Then
        Call GUARDAVALCOT(TRIM$(Label12), "TROQUELSTOCK", Text20)
    End If
    '
    'GUARDA LA FORMA
    If Option3.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMA", "1 - RECTANGULAR")    'RECTANGULAR
    If Option4.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMA", "2 - IRREGULAR")    'IRREGULAR
    If Option5.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMA", "3 - CIRCULAR / OVALO")    'CIRCULAR / OVALO
    If Option6.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMA", "4 - CORTE MEDIO")    'CORTE MEDIO
    '
    'GUARDA FORMA DE ENTREGA
    If Option25.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMAENTREGA", "1 - BOBINA")    'BOBINA
    If Option24.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMAENTREGA", "2 - PLANCHAS")    'PLANCHAS
    If Option26.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "FORMAENTREGA", "3 - INDISTINTO")    'INDISTINTO
    '
    'Call GUARDAVALCOT(TRIM$(Label12), "CAMBIOS", Text11)
    '
    Call GUARDAVALCOT(TRIM$(Label12), "IMPRCOLOR", Text15)
    '
    Call GUARDAVALCOT(TRIM$(Label12), "IMPRSTAMPING", Text13)
    '
    Call GUARDAVALCOT(TRIM$(Label12), "IMPRERELIEVE", Text16)
    '
    'GUARDA PORCENTAJE DE TINTA
    If Option12.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PORCENTAJETINTA", "1 - 10%")    '10%
    If Option8.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PORCENTAJETINTA", "2 - 30%")     '30%
    If Option11.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PORCENTAJETINTA", "3 - 60%")    '60%
    If Option7.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PORCENTAJETINTA", "4 - 80%")     '80%
    If Option13.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PORCENTAJETINTA", "5 - 100%")    '100%
    '
    If Check10.Value = 1 Then
        Call GUARDAVALCOT(TRIM$(Label12), "PANTONE", "SI")    'PANTONE
    Else
        Call GUARDAVALCOT(TRIM$(Label12), "PANTONE", "NO")    'PANTONE
    End If
    '
    'GRABA TERMINACION
    If Option31(0).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "1 - BARNIZ AGUA")    'BARNIZ AGUA
    If Option31(1).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "2 - BARNIZ UV")    'BARNIZ UV
    If Option31(2).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "3 - LAMINADO")    'LAMINADO
    If Option31(3).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "4 - SECTORIZADO")    'SECTORIZADO
    If Option31(4).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "5 - NINGUNO")    'NINGUNO
    If Option31(5).Value = True Then                                                          'OTRO
        'Call GUARDAVALCOT(TRIM$(Label12), "TERMINACION", "6 - " & Text17)    'OTRO
        ' SE UTILIZA ESTA FORMA PARA EVITAR EL TRIM DE LA FUNCION
        CONDI3$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'TERMINACION'"
        RR& = CantRegistros&("DETCOTIZA", CONDI3$)
        If RR& = 0 Then
            Call ABREPEDCLI
            sqlx$ = "SELECT * FROM DETCOTIZA"
            Dim DECOTERMTMP As ADODB.Recordset
            Set DECOTERMTMP = New ADODB.Recordset
            DECOTERMTMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
            With DECOTERMTMP
                .AddNew   ' YA CONVERTIDO
                ON ERROR RESUME NEXT
                !CodiEmpr = Codiemp%
                ON ERROR GOTO 0
                    !NroCoti = TRIM$(Label12)
                    !ClaValor = "TERMINACION"
                    !Valor1 = "6 - " & Text17
                .Update
             End With
        Else
           Call ACTUREGISTRO("DETCOTIZA", "Valor1", CONDI3$, "6 - " & Text17, RAF&, "NOMESS")
        End If
    End If
    '
    If Check20.Value = 1 Then
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAFRIO", "SI")       'FRIO
    Else
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAFRIO", "NO")       'FRIO
    End If
    '
    If Check22.Value = 1 Then
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAHUMEDAD", "SI")    'HUMEDAD
    Else
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAHUMEDAD", "NO")    'HUMEDAD
    End If
    '
    If Check21.Value = 1 Then
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAOTRO", "SI")       'OTRO
    Else
        Call GUARDAVALCOT(TRIM$(Label12), "SOPORTAOTRO", "NO")       'OTRO
    End If
    '
    'GUARDA APLICACION 1
    If Option40(0).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "APLICACION_1", "1 - MANUAL")  'MANUAL
    If Option40(1).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "APLICACION_1", "2 - AUTOMATICO")  'AUTOMATICO
    '
    'GUARDA APLICACION 2
    If Option32(0).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "APLICACION_2", "1 - IMPRESORA")  'IMPRESORA
    If Option32(1).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "APLICACION_2", "2 - DISPENSORA")  'DISPENSORA
    If Option32(2).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "APLICACION_2", "3 - UNIDADES")  'UNIDADES
    '
    'GRABA DATOS PRODUCTOS
    '
    'GUARDA REPITE
    If Option17.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "REPITE", "1 - REPETICION")  'REPETICION
    If Option16.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "REPITE", "2 - NUEVO")  'NUEVO
    '
    'GUARDA TIPO
    If Option1.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TIPO", "1 - FLEXOGRAFIA")      'FLEXOGRAFIA
    If Option2.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TIPO", "2 - DIGITAL")          'DIGITAL
    '
    'GUARDA PRODUCTO
    'Call GUARDAVALCOT(TRIM$(Label12), "PRODUCTO", Text4)  SE COMENTO PORQUE SE GRABA EN PRESUVEN
    '
    'GUARDA DESCRIPCION
    'Call GUARDAVALCOT(TRIM$(Label12), "DESCRIPCION", Text5)SE COMENTO PORQUE SE GRABA EN PRESUVEN
    '
    'GUARDA FAMILIA
    Call GUARDAVALCOT(TRIM$(Label12), "FAMILIA", Text35)
    '
    'GUARDA ADHESIVO
    Call GUARDAVALCOT(TRIM$(Label12), "MATERIALADHESIVO", Text6)
    '
    'GUARDA LINER
    Call GUARDAVALCOT(TRIM$(Label12), "MATERIALINER", Text7)
    '
    'GUARDA MUESTRA
    If Option30(0).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "MUESTRA", "1 - PRINT")  'PRINT
    If Option30(1).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "MUESTRA", "2 - MUESTRA")  'MUESTRA
    If Option30(2).Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "MUESTRA", "3 - S/DATOS")  'S/DATOS
    '
    'GUARDA TIPO IMPRESION PLUMA/FOTOCROMO/COD.BARRAS
    If Option28.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TIPOIMPRESION", "1 - PLUMA")  'PLUMA
    If Option27.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TIPOIMPRESION", "2 - FOTOCROMO")  'FOTOCROMO
    If Option29.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "TIPOIMPRESION", "3 - COD.BARRAS")  'COD.BARRAS
    '
    'DATOS COTIZACION
    '
    'GUARDA CANTIDADES  COTIZADASY PRECIOS
    Call GUARDAVALCOT(TRIM$(Label12), "CANTIDAD1", Text21(0))
    Call GUARDAVALCOT(TRIM$(Label12), "PRECIO1", Text24(0))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTIDAD2", Text21(1))
    Call GUARDAVALCOT(TRIM$(Label12), "PRECIO2", Text24(1))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTIDAD3", Text21(2))
    Call GUARDAVALCOT(TRIM$(Label12), "PRECIO3", Text24(2))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTIDAD4", Text21(3))
    Call GUARDAVALCOT(TRIM$(Label12), "PRECIO4", Text24(3))
    '
    'GUARDA CANTIDADES A COTIZAR Y OBSERVACIONES
    Call GUARDAVALCOT(TRIM$(Label12), "CANTACOTIZAR_1", Text22(0))
    Call GUARDAVALCOT(TRIM$(Label12), "OBSERVA_1", Text1(0))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTACOTIZAR_2", Text22(1))
    Call GUARDAVALCOT(TRIM$(Label12), "OBSERVA_2", Text1(1))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTACOTIZAR_3", Text22(2))
    Call GUARDAVALCOT(TRIM$(Label12), "OBSERVA_3", Text1(2))
    Call GUARDAVALCOT(TRIM$(Label12), "CANTACOTIZAR_4", Text22(3))
    Call GUARDAVALCOT(TRIM$(Label12), "OBSERVA_4", Text1(3))
    '
    'GUARDA DATOS COMPETENCIA
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM11", Text23(0))
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM12", Text25(0))
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM21", Text23(1))
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM22", Text25(1))
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM31", Text23(2))
    Call GUARDAVALCOT(TRIM$(Label12), "DATOCOM32", Text25(2))
    '
    'GASTOS DE IMPRESION
    Call GUARDAVALCOT(TRIM$(Label12), "GASTOSFOTO", "") 'Text31
    Call GUARDAVALCOT(TRIM$(Label12), "GASTOSTROQUEL", "") 'Text30)
    Call GUARDAVALCOT(TRIM$(Label12), "GASTOSSTAMPING", "") 'Text29)
    Call GUARDAVALCOT(TRIM$(Label12), "GASTOSRELIEVE", "") ' Text28)
    '
    Call GUARDAVALCOT(TRIM$(Label12), "PLAZOENTREGA", "") ' Text32)
    Call GUARDAVALCOT(TRIM$(Label12), "RECOMIENDA", "") 'Text33)
    '
    If Option33.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PREIMPRESION", "1 - PRECIO POR MILLAR")  '
    If Option10.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "PREIMPRESION", "2 - COTIZA POR SEPARADO")  'COD.BARRAS
    '
    If Option9.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "RESULTADO", "1 - PENDIENTE")  '
    If Option14.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "RESULTADO", "2 - GANADO")  'COD.BARRAS
    If Option15.Value = True Then Call GUARDAVALCOT(TRIM$(Label12), "RESULTADO", "3 - PERDIDO")  'COD.BARRAS
    '
    'Call GUARDAVALCOT(TRIM$(Label12), "MOTIVO", Text34)
    '
    'LOS CAMPOS DE CAMBIOS Y MOTIVO SE ACTUALIZAN POR FUERA DE LA FUNCION YA QUE LA MISMA REALIZA
    'UN TRIM DEL TEXT CORTANDO LOS MISMO EN LOS "enter"
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'MOTIVO'"
    RR& = CantRegistros&("DETCOTIZA", CONDI$)
    If RR& = 0 Then
        Call ABREPEDCLI
        sqlx$ = "SELECT * FROM DETCOTIZA"
        Dim DECOMOTITMP As ADODB.Recordset
        Set DECOMOTITMP = New ADODB.Recordset
        DECOMOTITMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With DECOMOTITMP
            .AddNew   ' YA CONVERTIDO
            ON ERROR RESUME NEXT
            !CodiEmpr = Codiemp%
            ON ERROR GOTO 0
                !NroCoti = TRIM$(Label12)
                !ClaValor = "MOTIVO"
                !Valor1 = Text34
            .Update
         End With
    Else
       Call ACTUREGISTRO("DETCOTIZA", "Valor1", CONDI$, Text34, RAF&, "NOMESS")
    End If
    
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'CAMBIOS'"
    RR& = CantRegistros&("DETCOTIZA", CONDI$)
    If RR& = 0 Then
        Call ABREPEDCLI
        sqlx$ = "SELECT * FROM DETCOTIZA"
        Dim DECOCAMBTMP As ADODB.Recordset
        Set DECOCAMBTMP = New ADODB.Recordset
        DECOCAMBTMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With DECOCAMBTMP
            .AddNew   ' YA CONVERTIDO
            ON ERROR RESUME NEXT
            !CodiEmpr = Codiemp%
            ON ERROR GOTO 0
                !NroCoti = TRIM$(Label12)
                !ClaValor = "CAMBIOS"
                !Valor1 = Text11
            .Update
         End With
    Else
       Call ACTUREGISTRO("DETCOTIZA", "Valor1", CONDI$, Text11, RAF&, "NOMESS")
    End If

99  Command5.Enabled = True
    
End Sub

Private Sub Command6_Click()
    '
   Call SELECHO("VENDEDOR")
   VEND$ = VALACT1$("VENDEDOR")
   If TRIM$(VEND$) <> "" Then
      Text2 = VEND$
      Label9 = VALACT2$("VENDEDOR")
   End If
   '
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    '
    Label4 = FECHATEX$(HOY(HO$))       'FECHA ACTUAL
    '
    Label12 = ""
    Call BLANFORMU
    '
    NUPROC$ = NUMECOTIZA(NUPROC$)
    '
    VDEF$ = String$(64, 0)
    FF% = HOY(HO$)
    Call REPLA(VDEF$, NUPROC$, 1, 16)
    Call REPLA(VDEF$, MKI$(FF%), 17, 2)
30  OBX$ = OBJPLA$("NUCOTIZA", VDEF$)
    '
    If OBX$ = "" Then GoTo 99
    '
    NUCOTIZA$ = TRIM$(Mid$(OBX$, 1, 16))
    FECC% = CVI(Mid$(OBX$, 17, 2))
    PROSPECT$ = TRIM$(Mid$(OBX$, 19, 12))
    NNCC% = CVI(Mid$(OBX$, 31, 2))
    MERCA% = CVI(Mid$(OBX$, 33, 2))
    AVANC = CVS(Mid$(OBX$, 35, 4))
    ANCHH = CVS(Mid$(OBX$, 39, 4))
    FORMM$ = TRIM$(Mid$(OBX$, 43, 12))
    FAMIL% = CVI(Mid$(OBX$, 55, 2))
    '
    Call REPLA(VDEF$, NUCOTIZA$, 1, 16)
    Call REPLA(VDEF$, MKI$(FECC%), 17, 2)
    Call REPLA(VDEF$, PROSPECT$, 19, 12)
    Call REPLA(VDEF$, MKI$(NNCC%), 31, 2)
    Call REPLA(VDEF$, MKI$(MERCA%), 33, 2)
    Call REPLA(VDEF$, MKS$(AVANC), 35, 4)
    Call REPLA(VDEF$, MKS$(ANCHH), 39, 4)
    Call REPLA(VDEF$, FORMM$, 43, 12)
    Call REPLA(VDEF$, MKI$(FAMIL%), 55, 2)
    '
    CCPP& = CantRegistros&("bsegcom", "Codi_Cli = '" & PROSPECT$ & "'")
    If CCPP& <= 0 And CODICLI$(NNCC%) = "" Then
        Call MENSERR(24, "Debe Seleccionar Cliente / Prospecto")
        GoTo 30
    End If
    If ECOARCH$("GRUCOCLI", MKI$(MERCA%)) = "" Then
        Call MENSERR(24, "Debe Seleccionar Mercado")
        GoTo 30
    End If
    If XVALO(AVANC) = 0 Then
        Call MENSERR(24, "Debe Ingresar Avance")
        GoTo 30
    End If
    If XVALO(ANCHH) = 0 Then
        Call MENSERR(24, "Debe Ingresar Ancho")
        GoTo 30
    End If
    If ECOARCH$("GRUCOVEN", MKI$(FAMIL%)) = "" Then
        Call MENSERR(24, "Debe Seleccionar Familia")
        GoTo 30
    End If
    '
    CCTT& = CantRegistros&("PRESUVEN", "NroCoti = '" & NUCOTIZA$ & "'")
    If CCTT& > 0 Then
        Call MENSERR(24, "Número de Cotización ya Registrado\ Flexoft Renumerará en Forma Automática")
        NUPROC$ = NUMECOTIZA(NUPROC$)
        Call REPLA(VDEF$, NUPROC$, 1, 16)
        Call REPLA(VDEF$, MKI$(FECC%), 17, 2)
        Call REPLA(VDEF$, PROSPECT$, 19, 12)
        Call REPLA(VDEF$, MKI$(NNCC%), 31, 2)
        Call REPLA(VDEF$, MKI$(MERCA%), 33, 2)
        Call REPLA(VDEF$, MKS$(AVANC), 35, 4)
        Call REPLA(VDEF$, MKS$(ANCHH), 39, 4)
        Call REPLA(VDEF$, FORMM$, 43, 12)
        Call REPLA(VDEF$, MKI$(FAMIL%), 55, 2)
    End If
    '
    Label12 = NUPROC$
    If NNCC% > 0 Then
        Label75 = NNCC%
        Label2 = FANTCLI$(NNCC%)
        If TRIM$(Label2) = "" Then Label2 = RASOCLI$(NNCC%)
        CONDI$ = "Nume_cli = '" & NNCC% & "'"
        Label3 = DOMICLI$(NNCC%)
        Label65 = LOCACLI$(NNCC%)
        Label67 = PAISCLI$(NNCC%)
        Text37 = DOMIENT$(NNCC%)
        Text38 = LOCAENT$(NNCC%)
    Else
        CONDI$ = "Codi_Cli = '" & TRIM$(PROSPECT$) & "'"
        Label2 = TRIM$(VALOBADA("BSEGCOM", "RASO_CLI", CONDI$, "NOMESS"))
        Label3 = VALOBADA("BSEGCOM", "DOMI_CLI", CONDI$)
        Label65 = VALOBADA("BSEGCOM", "LOCA_CLI", CONDI$)
        Label67 = VALOBADA("BSEGCOM", "PAIS_CLI", CONDI$)
        CONDI2$ = "Codi_Cli = '" & TRIM$(PROSPECT$) & "'"
        Label75 = TRIM$(VALOBADA("deudor12", "Nume_Cli", CONDI2$, "NOMESS"))
    End If
    Label74 = TRIM$(PROSPECT$)
    Text8 = TRIM$(VALOBADA("BSEGCOM", "Ctac_Cli", CONDI$, "NOMESS"))
    Text2 = XVALO(VALOBADA("BSEGCOM", "Vend_Cli", CONDI$, "NOMESS"))
    Text35 = FAMIL%
    Label4 = FECHATEX$(FECC%)
    Text10 = FORMATNUM$(AVANC, "F5.1")
    Text9 = FORMATNUM$(ANCHH, "F5.1")
    '
    Call ABREPEDCLI
    sqlx$ = "SELECT * FROM PRESUVEN"
    
    Dim COTIZATMP As ADODB.Recordset
    Set COTIZATMP = New ADODB.Recordset
    COTIZATMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    With COTIZATMP
        .AddNew   ' YA CONVERTIDO
        ON ERROR RESUME NEXT
        !CodiEmpr = Codiemp%
        ON ERROR GOTO 0
            !NroCoti = NUCOTIZA$
            !Codi_Cli = PROSPECT$
            !NroClie = NNCC%
            !FechEmis = FETEXCOR$(FECC%)
            !RasoClie = Label2
            !Vended = XVALO(Text2)
            !AteSen = TRIM$(Text8)
'            !Med_Avance = AVANC
'            !Med_Ancho = ANCHH
'            !Mercado = MERCA%
'            !Familia = FAMIL%
'            !FORMA = FORMM$
        .Update
    End With
    '
    'GENERA LISTA DE SELECCION INDICOTI
    XX$ = Space$(64)
    Call REPLA(XX$, TRIM$(Mid$(OBX$, 1, 16)), 1, 16)
    TIPP$ = ECOARCH$("GRUCOVEN", MKI$(FAMIL%))
    XYX$ = FECHATEX$(FF%) & " - " & TRIM$(Label2) & " - " & TRIM$(TIPP$)
    Call REPLA(XX$, XYX$, 17, 48)
    Call REGAPP("INDICOTI", XX$)
    Call FRESHECHO("INDICOTI")
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    '
    If DERACCE%("VERCOTI", "Visualizar Cotizaciones") > 0 Then
        Call SELECHO("INDICOTI")
        NUCO$ = VALACT1$("INDICOTI")
        If TRIM$(NUCO$) <> "" Then
          Label12 = TRIM$(NUCO$)
        End If
    End If
    '
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    STPCOT09.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    FF% = HOY(HO$)
    Label4 = FECHATEX(FF%)
    Label58 = FECHATEX(FF%)
    
    Call ABREPRESVEN
    '
End Sub



Private Sub Image3_Click()
    '
    If Label12.Caption = "" Then Exit Sub
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "'"
    '
    If DERACCE%("CIECOTI", "Cierre/Apertura de Cotizaciones") < 2 Then Exit Sub
    Image4.Visible = True
    Image3.Visible = False
    Call ACTUREGISTRO("PRESUVEN", "Status", CONDI$, 1, REG&, "NOMESS")
    '
End Sub

Private Sub Image4_Click()
    '
    If Label12.Caption = "" Then Exit Sub
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "'"
    '
    If DERACCE%("CIECOTI", "Cierre/Apertura de Cotizaciones") < 2 Then Exit Sub
    Image3.Visible = True
    Image4.Visible = False
    Call ACTUREGISTRO("PRESUVEN", "Status", CONDI$, 0, REG&, "NOMESS")
    '
End Sub

Private Sub Label12_Change()
    '
    Call BLANFORMU
    '
    CONDI$ = "NroCoti = '" & TRIM$(Label12) & "'"
    '
    ESTTD% = XVALO(VALOBADA("PRESUVEN", "Status", CONDI$, "NOMESS"))
    If ESTTD = 1 Then
        Image4.Visible = True
        Image3.Visible = False
    ElseIf ESTTD = 0 Then
        Image3.Visible = True
        Image4.Visible = False
    End If
    '
    Label74 = VALOBADA("PRESUVEN", "Codi_Cli", CONDI$, "NOMESS")
    NNCC% = XVALO(VALOBADA("PRESUVEN", "NroClie", CONDI$, "NOMESS"))
    If NNCC% > 0 Then
        Label75 = NNCC%
        Label3 = DOMICLI$(NNCC%)
        Label65 = LOCACLI$(NNCC%)
        Label67 = PAISCLI$(NNCC%)
    Else
        PROSPECT$ = VALOBADA("PRESUVEN", "CODI_CLI", CONDI$, "NOMESS")
        CONDI2$ = "Codi_Cli = '" & TRIM$(Label74) & "'"
        Label3 = VALOBADA("BSEGCOM", "DOMI_CLI", CONDI2$, "NOMESS")
        Label65 = VALOBADA("BSEGCOM", "LOCA_CLI", CONDI2$, "NOMESS")
        Label67 = VALOBADA("BSEGCOM", "PAIS_CLI", CONDI2$, "NOMESS")
        CONDI2$ = "Codi_Cli = '" & TRIM$(PROSPECT$) & "'"
        Label75 = TRIM$(VALOBADA("deudor12", "Nume_Cli", CONDI2$, "NOMESS"))
    End If
    '
    FFFF% = CVINT((VALOBADA("PRESUVEN", "FechEmis", CONDI$, "NOMESS")))
    Label4 = FECHATEX$(FFFF%)
    Label2 = VALOBADA("PRESUVEN", "RasoClie", CONDI$, "NOMESS")
    Text2 = VALOBADA("PRESUVEN", "Vended", CONDI$, "NOMESS")
    Text8 = VALOBADA("PRESUVEN", "AteSen", CONDI$, "NOMESS")
    Text3 = VALOBADA("PRESUVEN", "Reventa", CONDI$, "NOMESS")
    Text4 = CODEXT$(XVALO(VALOBADA("PRESUVEN", "CodiInt", CONDI$, "NOMESS")))
    Text5 = VALOBADA("PRESUVEN", "Especif", CONDI$, "NOMESS")
    Text19 = VALOBADA("PRESUVEN", "Antecedentes", CONDI$, "NOMESS")
    Text18 = VALOBADA("PRESUVEN", "Observa", CONDI$, "NOMESS")
    TIMAGEN$ = VALOBADA("PRESUVEN", "Imagen", CONDI$, "NOMESS")
    '
    'DATOS DEL CLIENTE
    '
    Text37 = VALEIDCOT(TRIM$(Label12), "LUGARENTREGA")
    Text38 = VALEIDCOT(TRIM$(Label12), "LOCAENTREGA")
    Text39 = VALEIDCOT(TRIM$(Label12), "PAISENTREGA")
    '
    'VENDEDOR LO LEVANTA DE PRESUVEN
    'Text2 = XVALO(VALEIDCOT(TRIM$(Label12), "VENDEDOR"))
    '
    'LEVANTA EL PRIVILEGIO DEL CLIENTE
    PRIV% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "CLIENTEPRIVILEGIA"), 1, 1))
    If PRIV% = 1 Then Option19.Value = True     'PRECIO
    If PRIV% = 2 Then Option20.Value = True     'CALIDAD
    If PRIV% = 3 Then Option18.Value = True     'ENTREGA RAPIDA
    '
    'LEVANTA LA CONDICIONES DE PAGO
    CCPAG% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "CONDPAGO"), 1, 1))
    If CCPAG% = 1 Then Option22.Value = True     'ANTICIPO Y CONTADO
    If CCPAG% = 2 Then Option23.Value = True     'CONTADO C/ENTREGA
    If CCPAG% = 3 Then Option21.Value = True     'Cta. Cte. Días
    '
    Label58 = VALEIDCOT(TRIM$(Label12), "FECHAENTREGA")
    '
    'CARACTERISTICAS DEL PRODUCTO
    '
    Text10 = TRIM$(FORMATNUM$(XVALO(VALEIDCOT(TRIM$(Label12), "MEDIDAAVANCE")), "F5.1"))
    Text9 = TRIM$(FORMATNUM$(XVALO(VALEIDCOT(TRIM$(Label12), "MEDIDAANCHO")), "F5.1"))
    '
    If VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO") = "NUEVO" Then
        Check1.Value = 1
    End If
    
'    If VALEIDCOT(TRIM$(Label12), "TROQUELSTOCK") <> "" Then
'        Text12 = VALEIDCOT(TRIM$(Label12), "TROQUELSTOCK")
'    End If
    If VALEIDCOT(TRIM$(Label12), "TROQUELSTOCK") <> "" Then
        Text20 = VALEIDCOT(TRIM$(Label12), "TROQUELSTOCK")
    End If
    '
    'LEVANTA LA FORMA
    FORMM% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "FORMA"), 1, 1))
    If FORMM% = 1 Then Option3.Value = True     'RECTANGULAR
    If FORMM% = 2 Then Option4.Value = True     'IRREGULAR
    If FORMM% = 3 Then Option5.Value = True     'CIRCULAR / OVALO
    If FORMM% = 4 Then Option6.Value = True     'CORTE MEDIO
    '
    'LEVANTA FORMA DE ENTREGA
    FORMENTR% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "FORMAENTREGA"), 1, 1))
    If FORMENTR% = 1 Then Option25.Value = True     'BOBINA
    If FORMENTR% = 2 Then Option24.Value = True     'PLANCHAS
    If FORMENTR% = 3 Then Option26.Value = True     'INDISTINTO
    '
    Text15 = VALEIDCOT(TRIM$(Label12), "IMPRCOLOR")
    '
    Text13 = VALEIDCOT(TRIM$(Label12), "IMPRSTAMPING")
    '
    Text16 = VALEIDCOT(TRIM$(Label12), "IMPRERELIEVE")
    '
    'LEVANTA PORCENTAJE DE TINTA
    TINTT% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "PORCENTAJETINTA"), 1, 1))
    If TINTT% = 1 Then Option12.Value = True    '10%
    If TINTT% = 2 Then Option8.Value = True    '30%
    If TINTT% = 3 Then Option11.Value = True    '60%
    If TINTT% = 4 Then Option7.Value = True    '80%
    If TINTT% = 5 Then Option13.Value = True    '100%
    '
    PANTO$ = VALEIDCOT(TRIM$(Label12), "PANTONE")
    If TRIM$(PANTO$) = "SI" Then
        Check10.Value = 1
    Else
        Check10.Value = 0
    End If
    '
    'LEVANTA TERMINACION
    TERMIN% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "TERMINACION"), 1, 1))
    If TERMIN% = 1 Then Option31(0).Value = True    'BARNIZ AGUA
    If TERMIN% = 2 Then Option31(1).Value = True    'BARNIZ UV
    If TERMIN% = 3 Then Option31(2).Value = True    'LAMIDO
    If TERMIN% = 4 Then Option31(3).Value = True    'SECTORIZADO
    If TERMIN% = 5 Then Option31(4).Value = True    'NINGUNO
    If TERMIN% = 6 Then
        Option31(5).Value = True                    'OTRO
        'Text17 = Mid$(VALEIDCOT(TRIM$(Label12), "TERMINACION"), 4)
        'SE UTILIZA VALOBADA PARA EVITAR TRIM DE LA FUNCION
        CONDI3$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'TERMINACION'"
        Text17 = Mid$(VALOBADA("DETCOTIZA", "Valor1", CONDI3$, "NOMESS"), 5)
    End If
    '
    SOPFRIO$ = VALEIDCOT(TRIM$(Label12), "SOPORTAFRIO")     'FRIO
    If TRIM$(SOPFRIO$) = "SI" Then
        Check20.Value = 1
    Else
        Check20.Value = 0
    End If
    '
    SOPHUM$ = VALEIDCOT(TRIM$(Label12), "SOPORTAHUMEDAD")   'HUMEDAD
    If TRIM$(SOPHUM$) = "SI" Then
        Check22.Value = 1
    Else
        Check22.Value = 0
    End If
    '
    SOPOTR$ = VALEIDCOT(TRIM$(Label12), "SOPORTAOTRO")      'OTRO
    If TRIM$(SOPOTR$) = "SI" Then
        Check21.Value = 1
    Else
        Check21.Value = 0
    End If
    '
    'LEVANTA APLICACION 1
    APLI1% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "APLICACION_1"), 1, 1))
    If APLI1% = 1 Then Option40(0).Value = True      'MANUAL
    If APLI1% = 2 Then Option40(1).Value = True      'AUTOMATICA
    '
    'LEVANTA APLICACION 2
    APLI2% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "APLICACION_2"), 1, 1))
    If APLI2% = 1 Then Option32(0).Value = True
    If APLI2% = 2 Then Option32(1).Value = True
    If APLI2% = 3 Then Option32(2).Value = True
    '
    'DATOS PRODUCTOS
    '
    'LEVANTA REPITE
    REPIT% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "REPITE"), 1, 1))
    If REPIT% = 1 Then Option17.Value = True             'REPETICION
    If REPIT% = 2 Then Option16.Value = True             'NUEVO
    '
    'LEVANTA TIPO
    TIPP% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "TIPO"), 1, 1))
    If TIPP% = 1 Then Option1.Value = True             'FLEXO
    If TIPP% = 2 Then Option2.Value = True             'DIGITAL
    '
    'LEVANTA PRODUCTO / LO LEVANTA DE PRESUVEN
'    Text4 = VALEIDCOT(TRIM$(Label12), "PRODUCTO")
'    '
'    Text5 = VALEIDCOT(TRIM$(Label12), "DESCRIPCION")
    '
    Text35 = VALEIDCOT(TRIM$(Label12), "FAMILIA")
    '
    Text6 = VALEIDCOT(TRIM$(Label12), "MATERIALADHESIVO")
    '
    Text7 = VALEIDCOT(TRIM$(Label12), "MATERIALINER")
    '
    'LEVANTA MUESTRA
    MUESTRR% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "MUESTRA"), 1, 1))
    If MUESTRR% = 1 Then Option30(0).Value = True
    If MUESTRR% = 2 Then Option30(1).Value = True
    If MUESTRR% = 3 Then Option30(2).Value = True
    '
    'LEVANTA TIP0 IMPRESION PLUMA/FOTOCROMO/COD.BARRAS
    TIPIMPP% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "TIPOIMPRESION"), 1, 1))
    If TIPIMPP% = 1 Then Option28.Value = True       'PLUMA
    If TIPIMPP% = 2 Then Option27.Value = True       'FOTOCROMO
    If TIPIMPP% = 3 Then Option29.Value = True       'COD.BARRAS
    '
    'LEVANTA DATOS DE COTIZACION
    '
    Text21(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTIDAD1"))
    Text24(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "PRECIO1"))
    Text21(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTIDAD2"))
    Text24(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "PRECIO2"))
    Text21(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTIDAD3"))
    Text24(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "PRECIO3"))
    Text21(3) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTIDAD4"))
    Text24(3) = TRIM$(VALEIDCOT(TRIM$(Label12), "PRECIO4"))
    '
    Text22(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTACOTIZAR_1"))
    Text1(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "OBSERVA_1"))
    Text22(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTACOTIZAR_2"))
    Text1(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "OBSERVA_2"))
    Text22(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTACOTIZAR_3"))
    Text1(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "OBSERVA_3"))
    Text22(3) = TRIM$(VALEIDCOT(TRIM$(Label12), "CANTACOTIZAR_4"))
    Text1(3) = TRIM$(VALEIDCOT(TRIM$(Label12), "OBSERVA_4"))
    '
    '
    Text23(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM11"))
    Text25(0) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM12"))
    Text23(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM21"))
    Text25(1) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM22"))
    Text23(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM31"))
    Text25(2) = TRIM$(VALEIDCOT(TRIM$(Label12), "DATOCOM32"))

    
    'GASTOS DE IMPRESION
    '
    Text31 = VALEIDCOT(TRIM$(Label12), "GASTOSFOTO")
    Text30 = VALEIDCOT(TRIM$(Label12), "GASTOSTROQUEL")
    Text29 = VALEIDCOT(TRIM$(Label12), "GASTOSSTAMPING")
    Text28 = VALEIDCOT(TRIM$(Label12), "GASTOSRELIEVE")
    '
    'PLAZO ENTREGA
    '
    Text32 = VALEIDCOT(TRIM$(Label12), "PLAZOENTREGA")
    Text33 = VALEIDCOT(TRIM$(Label12), "RECOMIENDA")
    '
    RESSULL% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "RESULTADO"), 1, 1))
    If RESSULL% = 1 Then Option9.Value = True
    If RESSULL% = 2 Then Option14.Value = True
    If RESSULL% = 3 Then Option15.Value = True
    '
    PREIMPR% = XVALO(Mid$(VALEIDCOT(TRIM$(Label12), "PREIMPRESION"), 1, 1))
    If PREIMPR% = 1 Then Option33.Value = True
    If PREIMPR% = 2 Then Option10.Value = True
    
    'Text11 = TRIM$(FORMATNUM$(XVALO(VALEIDCOT(TRIM$(Label12), "TROQUELNUEVO")), "F10.1"))
    'Text34 = VALEIDCOT(TRIM$(Label12), "MOTIVO")
    'SE UTILIZA VALOBADA PARA OBTENER MOTIVO Y CAMBIOS Y DE ESA MANERA EVITAR EL TRIM
    CONDI1$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'MOTIVO'"
    Text34 = VALOBADA("DETCOTIZA", "Valor1", CONDI1$, "NOMESS")
    CONDI2$ = "NroCoti = '" & TRIM$(Label12) & "' AND ClaValor = 'CAMBIOS'"
    Text11 = VALOBADA("DETCOTIZA", "Valor1", CONDI2$, "NOMESS")
    '
    Label16 = ""
    Image1.Picture = LoadPicture("")
    Call LEEFOTO
    
End Sub

Private Sub Label12_DblClick()
    Call Command8_Click
End Sub

Private Sub Label58_DblClick()
    Call Command14_Click
End Sub




Private Sub Option27_LostFocus()
    SSTab1.SetFocus
    SSTab1.TabIndex = 35
End Sub

Private Sub Option28_LostFocus()
    SSTab1.SetFocus
    SSTab1.TabIndex = 35
End Sub


Private Sub Option29_Click()
        SSTab1.SetFocus
        SSTab1.TabIndex = 35
End Sub

Private Sub Option31_Click(Index As Integer)
    If Option31(5).Value = False Then
        Text17 = ""
        Text17.Enabled = False
    Else
        Text17.Enabled = True
    End If
End Sub

Private Sub Text10_LostFocus()
    Text10 = FORMATNUM$(XVALO(Text10), "F5.1")
End Sub



Private Sub Text12_Click()
    Text20.SetFocus
End Sub

Private Sub Text18_LostFocus()
    SSTab1.SetFocus
    SSTab1.TabIndex = 59
End Sub

Private Sub Text19_LostFocus()
    SSTab1.SetFocus
    SSTab1.TabIndex = 18
End Sub

Private Sub Text2_Change()
   VENDE = XVALO(Text2.TEXT)
   If VENDE < 0 Or VENDE > 255 Then
      Text2.TEXT = "": Exit Sub
   End If
   Text2 = VENDE
   Label9.Caption = ECOARCH$("VENDEDOR", MKI$(Int(VENDE)))
End Sub

Private Sub Text2_DblClick()
    '
    Call Command6_Click
    '
End Sub


Private Sub Text2_LostFocus()
   VENDE = XVALO(Text2.TEXT)
   If VENDE < 0 Or VENDE > 255 Then
      Text2.TEXT = "": Exit Sub
   End If
   Label9.Caption = ECOARCH$("VENDEDOR", MKI$(Int(VENDE)))
   If Label9 = "" Then Text2 = ""
End Sub



Private Sub Text20_Change()
    '
    CONDI$ = "CODTROQUEL = '" & Text20 & "'"
    Text12 = TRIM$(VALOBADA("TROQUELES", "DESCRITROQUEL", CONDI$, "NOMESS"))
    '
End Sub

Private Sub Text20_DblClick()
    '
    Call Command15_Click
    '
End Sub

Private Sub Text21_LostFocus(Index As Integer)
    For I = 0 To 3
        Text21(I) = TRIM$(FORMATNUM$(XVALO(Text21(I)), "F8.0"))
    Next
End Sub


Private Sub Text22_Change(Index As Integer)
    For I = 0 To 3
        Text22(I) = TRIM$(FORMATNUM$(XVALO(Text22(I)), "F8.0"))
    Next
End Sub

Private Sub Text24_LostFocus(Index As Integer)
    For I = 0 To 3
        Text24(I) = TRIM$(FORMATNUM$(XVALO(Text24(I)), "F9.2"))
    Next
End Sub

Private Sub Text34_LostFocus()
    SSTab1.SetFocus
    SSTab1.TabIndex = 87
End Sub


Private Sub Text35_Change()
   VALONU = XVALO(Text35.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text35.TEXT = "": Exit Sub
   End If
   Label62.Caption = ECOARCH$("GRUCOVEN", MKI$(Int(VALONU)))
End Sub

Private Sub Text35_DblClick()
    Call Command12_Click
End Sub


Private Sub Text35_LostFocus()
   VALONU = XVALO(Text35.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text35.TEXT = "": Exit Sub
   End If
   Label62.Caption = ECOARCH$("GRUCOVEN", MKI$(Int(VALONU)))
   If Label62 = "" Then Text35 = ""
End Sub

Private Sub Text4_Change()
    '
    CCII% = CODINT%(TRIM$(Text4))
    Label17 = DESCRIT$(CCII%)
    '
End Sub

Private Sub Text4_DblClick()
    '
    Call Command3_Click
    '
End Sub

Sub BLANFORMU()
    
    TIMAGEN$ = ""
    Text2 = ""
    Text3 = ""
    Text8 = ""
    Label74 = ""
    Label75 = ""
    Label65 = ""
    Label67 = ""
    Label2 = ""
    Text37 = ""
    Text38 = ""
    Text39 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text17 = ""
    Text10 = ""
    Text9 = ""
    Text11 = ""
    Text20 = ""
    Text12 = ""
    Text15 = ""
    Text13 = ""
    Text16 = ""
    Text18 = ""
    Text19 = ""
    For I = 0 To 3
        Text1(I) = ""
        Text21(I) = ""
        Text24(I) = ""
        Text22(I) = ""
    Next I
    For I = 0 To 2
        Text25(I) = ""
        Text23(I) = ""
    Next I

    Text20 = ""
    Text26 = ""
    
    Text27 = ""
    Text31 = ""
    Text30 = ""
    Text29 = ""
    Text28 = ""
    Text32 = ""
    Text34 = ""
    Check20.Value = 0
    Check21.Value = 0
    Check22.Value = 0
    Option1.Value = False
    Option2.Value = False
    Option17.Value = False
    Option16.Value = False
    Option19.Value = False
    Option33.Value = False
    Option10.Value = False
    Option9.Value = False
    Option14.Value = False
    Option15.Value = False
    Option20.Value = False
    Option18.Value = False
    Option23.Value = False
    Option22.Value = False
    Option21.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Option25.Value = False
    Option24.Value = False
    Option26.Value = False
    Option40(0).Value = False
    Option40(1).Value = False
    Option32(0).Value = False
    Option32(1).Value = False
    Option32(2).Value = False
    Option13.Value = False
    Option12.Value = False
    Option8.Value = False
    Option11.Value = False
    Option7.Value = False
    Option31(0).Value = False
    Option31(1).Value = False
    Option31(2).Value = False
    Option31(3).Value = False
    Option31(4).Value = False
    Option31(5).Value = False
    Option27.Value = False
    Option28.Value = False
    Option29.Value = False
    Option30(0).Value = False
    Option30(1).Value = False
    Option30(2).Value = False
    Picture1.Picture = LoadPicture("")
    Label16 = ""

End Sub

Function NUMECOTIZA(NUPROCO$)
    '
    FF% = HOY(HO$)
    FEXX$ = FETEXCOR1$(FF%)
    '
    NUCOT$ = REPLACAR(Format(Date, "yy/mm/dd"), "/", "")
    '
    YYEAR = Year(Date): MMES = Month(Date) 'OBTIENE EL AÑO Y EL MES
    '
    Call ABREPEDCLI
    '
    Dim NUCOTITMP As ADODB.Recordset
    Set NUCOTITMP = New ADODB.Recordset
    sqlx$ = "SELECT TOP 1 NroCoti FROM PRESUVEN "
    sqlx$ = sqlx$ & " WHERE (DATEPART(YEAR,FECHEMIS)=" & YYEAR
    sqlx$ = sqlx$ & " AND DATEPART(MONTH,FECHEMIS)= " & MMES & ")"
    sqlx$ = sqlx$ & " ORDER BY NroCoti DESC"
    Set NUCOTITMP =  FLEXCONN.EXECUTE (FILTSQL$ (sqlx$))
    On Error Resume Next
    NUCOTITMP.MoveFirst
    If Err Then GoTo 20
    NUCO$ = Mid$(NUCOTITMP!NroCoti, 8, 4)
    NUCOTITMP.Close
    Set NUCOTITMP = Nothing
    '
20  On Error GoTo 0
      
    NUCOTI% = XVALO(NUCO$) + 1
    NUMECOT$ = TRIM$(Str$(NUCOTI%))
    While Len(NUMECOT$) < 4: NUMECOT$ = "0" + NUMECOT$: Wend
    NUMECOTIZA = NUCOT$ & "-" & NUMECOT$
    '
End Function

Private Sub Text9_LostFocus()
    Text9 = FORMATNUM$(XVALO(Text9), "F5.1")
End Sub
'
Sub LEEFOTO()
    '
    ARCHIVO$ = LUCOM$ & "IMGETIQ.IMG"   ' Label12
    Call SAVEFILE(ARCHIVO$, TIMAGEN$)
    Call MUESTRAPLANO(ARCHIVO$)
    On Error Resume Next
    Kill ARCHIVO$
    On Error GoTo 0
    '
End Sub

Sub MUESTRAPLANO(COPLANO$)
   '
   Screen.MousePointer = 11
   '
   Picture1.Width = 3615
   Picture1.Height = 3135
   Picture1.Visible = True
   Refresh
   '
   Image1.Picture = LoadPicture("")
   Picture1.Cls
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         On Error GoTo 90
         Image1.Picture = LoadPicture(COPLANO$)
         Picture1.Picture = LoadPicture("")
         Picture1.Refresh
         '
         HIMAGEN = Image1.Picture.Width
         VIMAGEN = Image1.Picture.Height
         
         HCUADRO = Picture1.Width
         VCUADRO = Picture1.Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1.Height = VCUADRO * REDUV / REDUC
         Picture1.Width = HCUADRO * REDUH / REDUC
         Picture1.Refresh
         '
         HCUADRO = Picture1.Width
         VCUADRO = Picture1.Height
         Picture1.Left = (Frame26.Width - HCUADRO) / 2
         Picture1.Top = (Frame26.Height - VCUADRO) / 2
         '
         Picture1.PaintPicture Image1.Picture, 0, 0, HCUADRO, VCUADRO
         Picture1.Visible = True
         Picture1.Refresh
      End If
   End If
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
99 End Sub


