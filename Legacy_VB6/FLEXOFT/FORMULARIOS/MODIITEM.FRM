VERSION 5.00
Begin VB.Form MODIITEM 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificacion de Atributos de Item de Stock "
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10150
      TabIndex        =   64
      Top             =   2415
      Width           =   175
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
      Height          =   285
      Left            =   7875
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   " "
      Top             =   2835
      Width           =   2460
   End
   Begin VB.TextBox Text25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9870
      TabIndex        =   65
      Text            =   " "
      Top             =   2415
      Width           =   360
   End
   Begin VB.CommandButton Command19 
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
      Height          =   855
      Left            =   5670
      Picture         =   "MODIITEM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Imprime Ficha de Artículo Activo"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   61
      Top             =   1995
      Width           =   1725
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8295
      TabIndex        =   5
      Top             =   1365
      Width           =   2040
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7875
      TabIndex        =   6
      Top             =   1890
      Width           =   2460
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Planos"
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
      Left            =   9450
      Picture         =   "MODIITEM.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Acceso a Planos , Fotos, Graficos"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4830
      TabIndex        =   55
      Top             =   3780
      Width           =   1200
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Tables"
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
      Left            =   4725
      Picture         =   "MODIITEM.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5880
      Width           =   855
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
      Height          =   360
      Left            =   10135
      TabIndex        =   53
      Top             =   4725
      Width           =   175
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
      Height          =   360
      Left            =   7455
      TabIndex        =   52
      Top             =   4725
      Width           =   175
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
      Height          =   360
      Left            =   2520
      TabIndex        =   51
      Top             =   4725
      Width           =   175
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
      Height          =   360
      Left            =   2205
      TabIndex        =   50
      Top             =   3780
      Width           =   175
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
      Height          =   360
      Left            =   6930
      TabIndex        =   49
      Top             =   2835
      Width           =   175
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
      Height          =   360
      Left            =   2100
      TabIndex        =   48
      Top             =   2835
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
      Height          =   360
      Left            =   2205
      TabIndex        =   47
      Top             =   1995
      Width           =   175
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
      Height          =   360
      Left            =   4095
      TabIndex        =   46
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rutas"
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
      Left            =   7560
      Picture         =   "MODIITEM.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Fórmula"
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
      Left            =   6615
      Picture         =   "MODIITEM.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command6 
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
      Height          =   855
      Left            =   3780
      Picture         =   "MODIITEM.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8925
      TabIndex        =   10
      Top             =   3675
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Text            =   " "
      Top             =   210
      Width           =   2565
   End
   Begin VB.CommandButton Command4 
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
      Height          =   855
      Left            =   2835
      Picture         =   "MODIITEM.frx":16D4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   855
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
      Height          =   855
      Left            =   9450
      Picture         =   "MODIITEM.frx":19DE
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
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
      Left            =   7560
      Picture         =   "MODIITEM.frx":1B28
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   210
      Width           =   855
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
      Height          =   855
      Left            =   8505
      Picture         =   "MODIITEM.frx":1E32
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Left            =   105
      TabIndex        =   27
      Top             =   5775
      Width           =   2535
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   210
         Picture         =   "MODIITEM.frx":213C
         Stretch         =   -1  'True
         Top             =   105
         Width           =   8925
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5985
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   " "
      Top             =   210
      Width           =   1200
   End
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
      Height          =   720
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "MODIITEM.frx":152A2
      Top             =   980
      Width           =   5625
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   3
      Text            =   "  "
      Top             =   1995
      Width           =   675
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
      Height          =   285
      Left            =   1575
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2415
      Width           =   4350
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   4
      Text            =   " "
      Top             =   2835
      Width           =   570
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6615
      TabIndex        =   7
      Text            =   " "
      Top             =   2835
      Width           =   360
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
      Height          =   285
      Left            =   1575
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   " "
      Top             =   3255
      Width           =   2250
   End
   Begin VB.TextBox Text9 
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
      Left            =   4095
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   3255
      Width           =   2985
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   12
      Text            =   " "
      Top             =   4725
      Width           =   990
   End
   Begin VB.TextBox Text11 
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
      Left            =   1575
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5250
      Width           =   4455
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      TabIndex        =   13
      Top             =   4725
      Width           =   1620
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7140
      TabIndex        =   14
      Text            =   " "
      Top             =   4725
      Width           =   360
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8925
      TabIndex        =   15
      Text            =   " "
      Top             =   4725
      Width           =   1220
   End
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
      Left            =   7140
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5250
      Width           =   3195
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   8
      Text            =   "  "
      Top             =   3780
      Width           =   675
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
      Height          =   285
      Left            =   1575
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8925
      TabIndex        =   9
      Top             =   3255
      Width           =   1410
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8925
      TabIndex        =   11
      Top             =   4095
      Width           =   1410
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Especif."
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
      Left            =   8505
      Picture         =   "MODIITEM.frx":152A4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descomposicion:"
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
      Left            =   7980
      TabIndex        =   67
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kilos / Título:"
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
      Left            =   2625
      TabIndex        =   62
      Top             =   2100
      Width           =   1500
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ta./ Tp:"
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
      Left            =   7350
      TabIndex        =   60
      Top             =   1470
      Width           =   870
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
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
      Left            =   6510
      TabIndex        =   59
      Top             =   1995
      Width           =   1290
   End
   Begin VB.Label Label17 
      Caption         =   "Label8"
      Height          =   120
      Left            =   4320
      TabIndex        =   57
      Top             =   3990
      Width           =   15
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.C.P.:"
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
      Left            =   3465
      TabIndex        =   56
      Top             =   3885
      Width           =   1290
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Máximo:"
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
      Left            =   7140
      TabIndex        =   42
      Top             =   3885
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      Left            =   585
      TabIndex        =   41
      Top             =   315
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   4725
      TabIndex        =   40
      Top             =   315
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   105
      TabIndex        =   39
      Top             =   980
      Width           =   1380
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Costo de Reposición:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2940
      TabIndex        =   38
      Top             =   4620
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Costeo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7665
      TabIndex        =   37
      Top             =   4620
      Width           =   1185
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Medida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   36
      Top             =   1890
      Width           =   1185
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimo:"
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
      Left            =   6510
      TabIndex        =   35
      Top             =   3465
      Width           =   2340
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   120
      Left            =   3900
      TabIndex        =   34
      Top             =   4935
      Width           =   15
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Económico:"
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
      Left            =   7035
      TabIndex        =   33
      Top             =   4200
      Width           =   1770
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Material:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   32
      Top             =   2730
      Width           =   1380
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Criterio de Reposicion."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4935
      TabIndex        =   31
      Top             =   2730
      Width           =   1590
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
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
      Left            =   5880
      TabIndex        =   30
      Top             =   4830
      Width           =   1185
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   210
      TabIndex        =   29
      Top             =   4830
      Width           =   1185
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Contable:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   28
      Top             =   3675
      Width           =   1290
   End
End
Attribute VB_Name = "MODIITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command16_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
         Text14.Text = VDEVU$
       Else
         Text14.Text = HO$
    End If
    Text14.SetFocus
End Sub

Private Sub Command17_Click()
    Call OPNOIN("")
End Sub

Private Sub Command18_Click()
    COD$ = TRIM$(Text1.Text)
    If COD$ <> "" And CODINT%(COD$) > 0 Then
        Command18.Enabled = False
        GRAFITEM.Label1.Caption = Text1.Text
        GRAFITEM.Label2.Caption = Text3.Text
        GRAFITEM.Label37.Caption = Text4.Text
        GRAFITEM.Label28.Caption = Text2.Text
        GRAFITEM.Show 1
        'Call CONECRUN("*MASTSPEC/" + COD$, "Especificaciones por Artículo")
        Command18.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Código de Artículo")
    End If
    '
    'If COD$ <> "" Then
    '    Command18.Enabled = False
    '    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '    If InStr(EPAC$, "NEUMANN") > O Then
    '      CIXI% = CODINT%(COD$)
    '      If CIXI% > 0 Then
    '        If CIXI% <= NUMAS% Then
    '          VDEF$ = REGBLAN$("INDIPLAN"): REPLANO& = 0
    '          RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
    '          If Len(RNC$) >= 4 Then
    '            REPLANO& = CVS(RNC$): If REPLANO& > ULTREG&("INDIPLAN") Then REPLANO& = 0
    '            If REPLANO& > 0 Then
    '              VDEF$ = REGLEIDO$("INDIPLAN", REPLANO&)
    '            End If
    '          End If
    '          Call REPLA(VDEF$, MKI$(CIXI%), 1, 2)
    '          If TRIM$(Mid$(VDEF$, 3, 16)) = "" Then
    '            Call REPLA(VDEF$, COD$, 3, 16)
    '          End If
    '          OBX$ = OBJPLA$("CARGAPLANO", VDEF$)
    '          If OBX$ <> "" Then
    '            If REPLANO& > 0 Then
    '                Call GRAREG("INDIPLAN", OBX$, REPLANO&)
    '              Else
    '                Call REGAPP("INDIPLAN", OBX$)
    '            End If
    '          End If
    '          Call CIERRARCH("INDIPLAN")
    '        End If
    '      End If
    '      Command18.Enabled = True
    '      Exit Sub
    '    End If
    '    '
    '    If EXISTE%("FICHACAL.EXE") > 0 Then
    '         Call CONECRUN("FICHACAL/" + COD$, "Ficha de Ingenieria de Calidad")
    '       Else
    '         CIIX% = CODINT%(COD$)
    '         Call MUESTRAPLANO(CIIX%)
    '    End If
    '    Command18.Enabled = True
    '  Else
    '    Call MENSERR(24, "Falta Ingresar Código de Artículo")
    'End If
End Sub

Private Sub Command19_Click()
    COD$ = TRIM$(Text1.Text)
    If COD$ <> "" Then
        PrintForm
      Else
        Call MENSERR(24, "Imposible Imprimir Ficha de Artículo.\Ingrese Previamente el Código de Producto.")
    End If
End Sub

Private Sub Command20_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACOMPO")
    If VALACT1$("TACOMPO") <> "" Then
      Text25.Text = VALACT1$("TACOMPO")
      Text25.Refresh
      Text26.Text = VALACT2$("TACOMPO")
      Text26.Refresh
    End If
End Sub

Private Sub Command6_Click()
    'Call OPNOIN("")
    OPGRADOSI.Show 1
End Sub

Private Sub Command7_Click()
    COD$ = TRIM$(Text1.Text)
    If COD$ <> "" Then
        Call CIERRARCH("*.*")
        Call CONECRUN("FORMULAS/" + COD$, "Fórmulas y Estructuras de Producto")
      Else
        Call MENSERR(24, "Imposible Acceder a Fórmulas y Estructuras.\Ingrese Previamente el Código de Producto.")
    End If
End Sub

Private Sub Command8_Click()
    COD$ = TRIM$(Text1.Text)
    If COD$ <> "" Then
        Call CIERRARCH("*.*")
        Call CONECRUN("HORUTAS/" + COD$, "Secuencia de Operaciones por Producto")
      Else
        Call MENSERR(24, "Imposible Acceder a Secuencia de Operaciones.\Ingrese Previamente el Código de Producto.")
    End If

End Sub

Private Sub Text1_Change()
    VALICOD% = 0
    MODIFI% = 0
    CODX$ = TRIM$(Text1.Text)
    If CODINT%(CODX$) > 0 Then
         COD$ = Text1.Text
         Call CARDATIT(COD$)
         MODIITEM.Refresh
         MODIFI% = 0
         Text3.SetFocus
    End If
End Sub

Private Sub Text1_DBLCLICK()
    If MODIFI% > 0 Then
      If TRIM$(Text1.Text) <> "" Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            command1.SetFocus
            Exit Sub
        End If
      End If
    End If
    '
    VALICOD% = 0
    COD$ = ""
    Call SELECHO("MASTER")
End Sub

Private Sub Command9_Click()
    If MODIFI% > 0 Then
      If TRIM$(Text1.Text) <> "" Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            command1.SetFocus
            Exit Sub
        End If
      End If
    End If
    '
    VALICOD% = 0
    COD$ = ""
    Call SELECHO("MASTER")
    Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
    If VALICOD% < 0 Then Exit Sub
    MODIITEM.Refresh
    COD$ = VALACT1$("MASTER")
    If CODINT%(COD$) > 0 Then
         Text1.Text = COD$
         Call CARDATIT(COD$)
         MODIITEM.Refresh
         MODIFI% = 0
         Text3.SetFocus
    End If
    If TRIM$(Text1.Text) = "" Then
        MODIFI% = 0
    End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VALICOD% = 0
        COD$ = TRIM$(Text1.Text)
        Text3.SetFocus
    End If
End Sub


Private Sub Command1_Click()
    '
    COD$ = AJUSTI$(Text1.Text, 16)
    If TRIM$(COD$) = "" Then
        Call MENSERR(24, "Código no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    If CODINT%(COD$) < 1 Then
        Call MENSERR(24, "Codigo Inexistente en Maestro")
        Text1.SetFocus
        Exit Sub
    End If
    '
    DE$ = TRIM$(Text3.Text)
    If DE$ = "" Then
        Call MENSERR(24, "Falta Descripción del Item")
        Text3.SetFocus
        Exit Sub
    End If
    '
    VALOR = Int(Val(Text6.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 15
    TIMA% = Val(Text6.Text)
    If TIMA% > 0 Then
        If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
15          TIMA% = 0
            Text6.Text = ""
            Text7.Text = ""
            Call MENSERR(24, "Tipo de Material no Válido")
            Text6.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text8.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 25
    CRIT% = Val(Text8.Text)
    If CRIT% > 0 Then
        If ECOARCH$("TACRIRE", Chr$(CRIT%)) = "" Then
25          CRIT% = 0
            Text8.Text = ""
            Text9.Text = ""
            Call MENSERR(24, "Criterio de Reposición no Válido")
            Text8.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text13.Text))
    If VALOR < 0 Or VALOR > 255 Then GoTo 35
    MONEI% = Val(Text13.Text)
    If MONEI% > 0 Then
        If ECOARCH$("TAMONED", Chr$(MONEI%)) = "" Then
35          MONEI% = 1
            Text13.Text = "1"
            Text15.Text = ECOARCH$("TAMONED", Chr$(1))
            Call MENSERR(24, "Moneda de Costeo no Válida")
            Text13.SetFocus
            Exit Sub
        End If
    End If
    '
    STMIN = Val(Text18.Text)
    If STMIN <= 30000 Then
          STMI& = STMIN
        ElseIf STMIN <= 3000000 Then
          STMI& = (-1) * STMIN / 100
        Else
          Call MENSERR(24, "Stock Mínimo fuera de Rango")
          Text18.Text = ""
          Text18.SetFocus
          Exit Sub
    End If
    STMII% = STMI&
    '
    LOREP = Val(Text19.Text)
    If LOREP <= 30000 Then
          LORE& = LOREP
        ElseIf LOREP <= 3000000 Then
          LORE& = (-1) * LOREP / 100
        Else
          Call MENSERR(24, "Lote Económico fuera de Rango")
          Text19.Text = ""
          Text19.SetFocus
          Exit Sub
    End If
    LOREI% = LORE&
    '
    LOTEMA = Val(Text20.Text)
    If LOTEMA <= 30000 Then
          LOMA& = LOTEMA
        ElseIf LOTEMA <= 3000000 Then
          LOMA& = (-1) * LOTEMA / 100
        Else
          Call MENSERR(24, "Lote de Reposición fuera de Rango")
          Text20.Text = ""
          Text20.SetFocus
          Exit Sub
    End If
    LOTEMAI% = LOMA&
    '
    UNID$ = Text4.Text
    COU = Val(Text12.Text)
    FECO% = FECHANUM(Text14.Text)
    FCP% = Val(Text21.Text)
    '
    VALOR = Int(Val(Text10.Text))
    If VALOR < 0 Or VALOR > 32767 Then GoTo 85
    NPRO% = Val(Text10.Text)
    If NPRO% > 0 Then
        If ECOARCH$("PROVED1", MKI$(NPRO%)) = "" Then
85          NPRO% = 0
            Text10.Text = ""
            Text11.Text = ""
            Call MENSERR(24, "Proveedor Inexistente")
            Text10.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text16.Text))
    If VALOR < 0 Or VALOR > 32767 Then GoTo 95
    GRUCO% = Val(Text16.Text)
    If GRUCO% > 0 Then
        If ECOARCH$("GRUCOVEN", MKI$(GRUCO%)) = "" Then
95          GRUCO% = 0
            Text16.Text = ""
            Text17.Text = ""
            Call MENSERR(24, "Grupo Contable de Venta no Válido")
            Text16.SetFocus
            Exit Sub
        End If
    End If
    '
    VALOR = Int(Val(Text25.Text))
    If VALOR < 0 Or VALOR > 1 Then GoTo 105
    CCOMPO% = Val(Text25.Text)
    If CCOMPO% > 0 Then
        If ECOARCH$("TACOMPO", Chr$(CCOMPO%)) = "" Then
105         CCOMPO% = 0
            Text25.Text = ""
            Text26.Text = ""
            Call MENSERR(24, "Criterio de Descomposición no Válido")
            Text25.SetFocus
            Exit Sub
        End If
    End If
    '
    TCOLOR$ = TRIM$(UCase$(Left$(Text22.Text, 31)))
    TTALLE$ = TRIM$(UCase$(Left$(Text23.Text, 15)))
    TTITPE$ = TRIM$(UCase$(Text24.Text))
    '
    Call MODIMAS
    MODIFI% = 0
    Text1.Text = ""
    Text1.SetFocus
    '
End Sub

Private Sub Command2_Click()
    Call CARDATIT(COD$)
    MODIFI% = 0
    Text3.SetFocus
End Sub

Private Sub Command3_Click()
    If TRIM$(Text1.Text) <> "" Then
        If MODIFI% > 0 Then
            RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
            If RESPU <> 6 Then
                command1.SetFocus
                Exit Sub
            End If
        End If
        '
10      Call CLEARMODI
        VALICOD% = (-1)
        MODIFI% = 0
        Text1.SetFocus
        Exit Sub
    End If
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("AMASTO")
    '
    If AGREPRECIO% = 1 Then
      Call COMUNI("Recuerde Ordenar Oportunamente\la Lista de Precios Principal.")
      AGREPRECIO% = 0
    End If
    '
    Call CLEARMODI
    VALICOD% = (-1)
    
    MODIITEM.Hide
    Call FRESHALL
End Sub

Private Sub Command4_Click()
    Call HELPONLINE("AMASTO")
End Sub

Private Sub Command5_Click()
    COD$ = TRIM$(Text1.Text)
    If COD$ <> "" And CODINT%(COD$) > 0 Then
        Command5.Enabled = False
        ESPECITEM.Label1.Caption = Text1.Text
        ESPECITEM.Label2.Caption = Text3.Text
        ESPECITEM.Label37.Caption = Text4.Text
        ESPECITEM.Label28.Caption = Text2.Text
        ESPECITEM.Show 1
        'Call CONECRUN("*MASTSPEC/" + COD$, "Especificaciones por Artículo")
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Código de Artículo")
    End If
End Sub

Private Sub Text10_Change()
    MODIFI% = 1
End Sub

Private Sub Text10_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("PROVED1")
    If VALACT1$("PROVED1") <> "" Then
       Text10.Text = VALACT1$("PROVED1")
       Text10.Refresh
       Text11.Text = VALACT2$("PROVED1")
       Text11.Refresh
    End If
End Sub

Private Sub Command14_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("PROVED1")
    If VALACT1$("PROVED1") <> "" Then
       Text10.Text = VALACT1$("PROVED1")
       Text10.Refresh
       Text11.Text = VALACT2$("PROVED1")
       Text11.Refresh
    End If
End Sub

Private Sub Text11_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text10.SetFocus
End Sub

Private Sub Text12_Change()
    MODIFI% = 1
End Sub

Private Sub Text13_Change()
    MODIFI% = 1
End Sub

Private Sub Text14_Change()
    MODIFI% = 1
End Sub

Private Sub Text14_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
         Text14.Text = VDEVU$
       Else
         Text14.Text = HO$
    End If
End Sub

Private Sub Text15_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text13.SetFocus
End Sub

Private Sub Text16_Change()
    MODIFI% = 1
End Sub

Private Sub Text16_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("GRUCOVEN")
    If VALACT1$("GRUCOVEN") <> "" Then
       Text16.Text = VALACT1$("GRUCOVEN")
       Text16.Refresh
       Text17.Text = VALACT2$("GRUCOVEN")
       Text17.Refresh
    End If
End Sub

Private Sub Command13_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("GRUCOVEN")
    If VALACT1$("GRUCOVEN") <> "" Then
       Text16.Text = VALACT1$("GRUCOVEN")
       Text16.Refresh
       Text17.Text = VALACT2$("GRUCOVEN")
       Text17.Refresh
    End If
End Sub

Private Sub Text17_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text16.SetFocus
End Sub

Private Sub Text18_Change()
    MODIFI% = 1
End Sub

Private Sub Text19_Change()
    If Val(Text19.Text) > Val(Text20.Text) Then
        Text20.Text = Text19.Text
    End If
    MODIFI% = 1
End Sub

Private Sub Text2_GotFocus()
    Text3.SetFocus
End Sub

Private Sub Text20_Change()
    MODIFI% = 1
End Sub
Private Sub Text21_Change()
    MODIFI% = 1
End Sub


Private Sub Text22_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text22_DblClick()
   Call SELECHO("TACOLOR")
   Text22.Text = VALACT1$("TACOLOR")
End Sub

Private Sub Text23_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text23_DblClick()
   Call SELECHO("TATALLES")
   Text23.Text = VALACT1$("TATALLES")
End Sub

Private Sub Text24_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text25_DblClick()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACOMPO")
    If VALACT1$("TACOMPO") <> "" Then
      Text25.Text = VALACT1$("TACOMPO")
      Text25.Refresh
      Text26.Text = VALACT2$("TACOMPO")
      Text26.Refresh
    End If
End Sub

Private Sub Text26_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text25.SetFocus
End Sub

Private Sub Text3_Change()
    MODIFI% = 1
End Sub

Private Sub Text3_GotFocus()
    If TRIM$(COD$) <> TRIM$(Text1.Text) Then
        Text1.Text = COD$
    End If
    If VALICOD% <> 1 Then
        COD$ = TRIM$(Text1.Text)
        If CODINT%(COD$) > 0 Then
            Call CARDATIT(COD$)
            Exit Sub
        End If
        Text1.SetFocus
    End If
End Sub

Private Sub Text4_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text4_Change()
    MODIFI% = 1
End Sub

Private Sub Text6_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text6_Change()
    MODIFI% = 1
End Sub

Private Sub Text8_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text12_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text21_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text13_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text14_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text16_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text18_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text19_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text10_Click()
    If VALICOD% <> 1 Then Text1.SetFocus
End Sub

Private Sub Text4_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text4.Text = ""
    Call SELECHO("UMEDIDA")
End Sub

Private Sub command10_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text4.Text = ""
    Call SELECHO("UMEDIDA")
    Text4.SetFocus
End Sub

Private Sub Text4_GOTFOCUS()
    If Text4.Text = "" Then
        Text4.Text = VALACT1$("UMEDIDA")
        Text4.Refresh
        Text5.Text = VALACT2$("UMEDIDA")
        Text5.Refresh
    End If
End Sub

Private Sub Text5_click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text4.SetFocus
End Sub

Private Sub Text6_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TATIMAT")
    If VALACT1$("TATIMAT") <> "" Then
       Text6.Text = VALACT1$("TATIMAT")
       Text6.Refresh
       Text7.Text = VALACT2$("TATIMAT")
       Text7.Refresh
    End If
End Sub

Private Sub Command11_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TATIMAT")
    If VALACT1$("TATIMAT") <> "" Then
       Text6.Text = VALACT1$("TATIMAT")
       Text6.Refresh
       Text7.Text = VALACT2$("TATIMAT")
       Text7.Refresh
    End If
End Sub

Private Sub Text7_click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text6.SetFocus
End Sub

Private Sub Text8_Change()
    MODIFI% = 1
End Sub

Private Sub Text8_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACRIRE")
    If VALACT1$("TACRIRE") <> "" Then
      Text8.Text = VALACT1$("TACRIRE")
      Text8.Refresh
      Text9.Text = VALACT2$("TACRIRE")
      Text9.Refresh
    End If
End Sub

Private Sub Command12_Click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TACRIRE")
    If VALACT1$("TACRIRE") <> "" Then
      Text8.Text = VALACT1$("TACRIRE")
      Text8.Refresh
      Text9.Text = VALACT2$("TACRIRE")
      Text9.Refresh
    End If
End Sub

Private Sub Text13_DBLCLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TAMONED")
    If VALACT1$("TAMONED") <> "" Then
       Text13.Text = VALACT1$("TAMONED")
       Text13.Refresh
       Text15.Text = VALACT2$("TAMONED")
       Text15.Refresh
    End If
End Sub

Private Sub COMMAND15_CLICK()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Call SELECHO("TAMONED")
    If VALACT1$("TAMONED") <> "" Then
       Text13.Text = VALACT1$("TAMONED")
       Text13.Refresh
       Text15.Text = VALACT2$("TAMONED")
       Text15.Refresh
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub
Private Sub Text9_click()
    If VALICOD% <> 1 Then Text1.SetFocus: Exit Sub
    Text8.SetFocus
End Sub

