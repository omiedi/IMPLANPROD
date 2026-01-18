VERSION 5.00
Begin VB.Form TRAZORF04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame9"
      Height          =   540
      Left            =   2940
      TabIndex        =   92
      Top             =   4410
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label34 
         Height          =   225
         Left            =   5775
         TabIndex        =   98
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label20 
         Height          =   225
         Left            =   4725
         TabIndex        =   97
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label17 
         Height          =   225
         Left            =   3570
         TabIndex        =   96
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label21 
         Height          =   225
         Left            =   2415
         TabIndex        =   95
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label33 
         Height          =   225
         Left            =   735
         TabIndex        =   94
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label Label8 
         Height          =   225
         Left            =   105
         TabIndex        =   93
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "# "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   6195
      TabIndex        =   89
      Top             =   210
      Width           =   1065
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
         Left            =   210
         TabIndex        =   9
         Top             =   1050
         Width           =   700
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
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   525
         Width           =   700
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -840
         TabIndex        =   91
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -840
         TabIndex        =   90
         Top             =   315
         Width           =   1335
      End
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
      Left            =   6510
      TabIndex        =   7
      Top             =   1785
      Width           =   175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   6090
      TabIndex        =   6
      Top             =   1785
      Width           =   435
   End
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
      Left            =   1785
      TabIndex        =   5
      Top             =   1785
      Width           =   175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   1365
      TabIndex        =   4
      Top             =   1785
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   3
      Top             =   1365
      Width           =   960
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5775
      TabIndex        =   1
      Top             =   315
      Width           =   175
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
      Left            =   4305
      TabIndex        =   2
      Top             =   735
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   7455
      TabIndex        =   74
      Top             =   210
      Width           =   2115
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   105
         TabIndex        =   75
         Top             =   315
         Width           =   1870
      End
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
      Left            =   3150
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   51
      Top             =   2205
      Width           =   9360
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   8900
         TabIndex        =   88
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   7570
         TabIndex        =   87
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   6250
         TabIndex        =   83
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   1850
         TabIndex        =   82
         Top             =   450
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   630
         Width           =   9000
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo                   Descripción                                                               Consumo            Qt.Trz."
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
         Left            =   210
         TabIndex        =   56
         Top             =   375
         Width           =   9000
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8500
      Left            =   9765
      ScaleHeight     =   8445
      ScaleWidth      =   1320
      TabIndex        =   44
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRAZORF04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   4515
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
         Picture         =   "TRAZORF04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3780
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   1110
         Left            =   105
         Picture         =   "TRAZORF04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Guardar Cierre de Orden de Fabricacion"
         Top             =   5775
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
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
         Picture         =   "TRAZORF04.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3045
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "TRAZORF04.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command16 
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
         Left            =   105
         Picture         =   "TRAZORF04.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Impresión de Registro de Trazabilidad"
         Top             =   2310
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   45
         Top             =   7035
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   12000
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "TRAZORF04.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRAZORF04.frx":14BE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7395
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Trazabilidad De Componentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   210
      TabIndex        =   47
      Top             =   4620
      Width           =   9360
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
         Height          =   750
         Left            =   8295
         Picture         =   "TRAZORF04.frx":33E0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   2310
         Width           =   855
      End
      Begin VB.CommandButton Command12 
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
         Height          =   250
         Left            =   8295
         TabIndex        =   37
         Top             =   1785
         Width           =   840
      End
      Begin VB.CommandButton Command11 
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
         Height          =   250
         Left            =   8295
         TabIndex        =   36
         Top             =   1470
         Width           =   840
      End
      Begin VB.TextBox Text4 
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
         Left            =   5460
         TabIndex        =   84
         Top             =   420
         Width           =   960
      End
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
         Height          =   250
         Index           =   5
         Left            =   3150
         TabIndex        =   31
         Top             =   2820
         Width           =   175
      End
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
         Height          =   250
         Index           =   4
         Left            =   3150
         TabIndex        =   27
         Top             =   2550
         Width           =   175
      End
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
         Height          =   250
         Index           =   3
         Left            =   3150
         TabIndex        =   23
         Top             =   2280
         Width           =   175
      End
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
         Height          =   250
         Index           =   2
         Left            =   3150
         TabIndex        =   19
         Top             =   2010
         Width           =   175
      End
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
         Height          =   250
         Index           =   1
         Left            =   3150
         TabIndex        =   15
         Top             =   1740
         Width           =   175
      End
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
         Height          =   250
         Index           =   0
         Left            =   3150
         TabIndex        =   11
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   5
         Left            =   3675
         TabIndex        =   32
         Top             =   2820
         Width           =   1065
      End
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
         Height          =   250
         Index           =   5
         Left            =   5040
         TabIndex        =   33
         Top             =   2820
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   6720
         TabIndex        =   34
         Top             =   2820
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   4
         Left            =   3675
         TabIndex        =   28
         Top             =   2550
         Width           =   1065
      End
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
         Height          =   250
         Index           =   4
         Left            =   5040
         TabIndex        =   29
         Top             =   2550
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6720
         TabIndex        =   30
         Top             =   2550
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   3
         Left            =   3675
         TabIndex        =   24
         Top             =   2280
         Width           =   1065
      End
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
         Height          =   250
         Index           =   3
         Left            =   5040
         TabIndex        =   25
         Top             =   2280
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6720
         TabIndex        =   26
         Top             =   2280
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   2
         Left            =   3675
         TabIndex        =   20
         Top             =   2010
         Width           =   1065
      End
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
         Height          =   250
         Index           =   2
         Left            =   5040
         TabIndex        =   21
         Top             =   2010
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6720
         TabIndex        =   22
         Top             =   2010
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   1
         Left            =   3675
         TabIndex        =   16
         Top             =   1740
         Width           =   1065
      End
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
         Height          =   250
         Index           =   1
         Left            =   5040
         TabIndex        =   17
         Top             =   1740
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6720
         TabIndex        =   18
         Top             =   1740
         Width           =   1380
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Index           =   0
         Left            =   3675
         TabIndex        =   12
         Top             =   1470
         Width           =   1065
      End
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
         Height          =   250
         Index           =   0
         Left            =   5040
         TabIndex        =   13
         Top             =   1470
         Width           =   1380
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
         Height          =   250
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   6720
         TabIndex        =   14
         Top             =   1470
         Width           =   1380
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   8190
         TabIndex        =   86
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Q.Traz:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   85
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   5
         Left            =   1470
         TabIndex        =   73
         Top             =   2820
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   4
         Left            =   1470
         TabIndex        =   72
         Top             =   2550
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   250
         Index           =   3
         Left            =   1470
         TabIndex        =   71
         Top             =   2280
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   250
         Index           =   2
         Left            =   1470
         TabIndex        =   70
         Top             =   2010
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   250
         Index           =   1
         Left            =   1470
         TabIndex        =   69
         Top             =   1740
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL / PLC Fin:"
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
         Left            =   5775
         TabIndex        =   68
         Top             =   1260
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL / PLC Com.:"
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
         Left            =   4200
         TabIndex        =   67
         Top             =   1260
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
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
         Left            =   1890
         TabIndex        =   66
         Top             =   1260
         Width           =   1965
      End
      Begin VB.Label Label29 
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
         Height          =   300
         Left            =   4095
         TabIndex        =   62
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   1470
         TabIndex        =   59
         ToolTipText     =   "Descripción"
         Top             =   735
         Width           =   7680
      End
      Begin VB.Label Label27 
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
         Left            =   1470
         TabIndex        =   58
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   105
         TabIndex        =   57
         Top             =   500
         Width           =   1335
      End
      Begin VB.Label Label23 
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
         Height          =   250
         Index           =   0
         Left            =   1470
         TabIndex        =   52
         Top             =   1470
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.Rec.:"
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
         Left            =   -210
         TabIndex        =   48
         Top             =   1260
         Width           =   1965
      End
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Consumo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4830
      TabIndex        =   81
      Top             =   1845
      Width           =   1125
   End
   Begin VB.Label Label11 
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
      Left            =   6825
      TabIndex        =   80
      Top             =   1785
      Width           =   2730
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep.Entra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   79
      Top             =   1850
      Width           =   1125
   End
   Begin VB.Label Label9 
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
      Left            =   2100
      TabIndex        =   78
      Top             =   1785
      Width           =   2595
   End
   Begin VB.Label Label7 
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
      Height          =   300
      Left            =   4200
      TabIndex        =   77
      Top             =   1155
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2940
      TabIndex        =   76
      Top             =   800
      Width           =   1335
   End
   Begin VB.Label Label32 
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
      Left            =   1365
      TabIndex        =   65
      ToolTipText     =   "Doble-Click para Refrescar"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label31 
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
      Height          =   300
      Left            =   105
      TabIndex        =   64
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Height          =   495
      Left            =   1365
      TabIndex        =   63
      Top             =   1155
      Width           =   3165
   End
   Begin VB.Label Label28 
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
      Left            =   1365
      TabIndex        =   61
      Top             =   735
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -315
      TabIndex        =   60
      Top             =   800
      Width           =   1545
   End
   Begin VB.Label Label14 
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
      Height          =   300
      Left            =   3465
      TabIndex        =   55
      Top             =   380
      Width           =   1335
   End
   Begin VB.Label Label24 
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
      Left            =   4830
      TabIndex        =   54
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Fabr. Nº:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   53
      Top             =   380
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   1470
      TabIndex        =   50
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   1470
      TabIndex        =   49
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "TRAZORF04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset

Private Sub command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command19_Click()

   Command19.Enabled = False
   Call INFOFA04.Command19_Click
   Command19.Enabled = True
   
End Sub

Private Sub Command26_Click()

   Command26.Enabled = False
   Text10 = ""
   Label4 = ""
   MODIPAMOD04.Show 1
   Command26.Enabled = True
   
End Sub

Private Sub Command29_Click()
   '
   Call OPNOIN("Ayuda FLEXOFT en Línea")
   '
End Sub

Private Sub Command10_Click()

    Command10.Enabled = False
    If List1.ListIndex >= 0 Then
       FIKARDEX.Text1 = Left$(List1.TEXT, 16)
    End If
    FIKARDEX.Show 1
    Command10.Enabled = True
    
End Sub


Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
       CICOMPO% = CODINT%(VDEVU$)
       CACONSU = 0
       TTYY$ = Space$(128)
       Call REPLA(TTYY$, VDEVU$, 1, 16)
       Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
       Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
       On Error Resume Next
       USOUNI = CACONSU / XVALO(Text1)
       On Error GoTo 0
       Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
       List1.AddItem TTYY$
       List1.ListIndex = List1.ListCount - 1
       Text4 = ""
       Text4.SetFocus
    End If
    '
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       Call SELECHO("MASTER")
       VDEVU$ = VALACT1$("MASTER")
       If VDEVU$ <> "" Then
          CICOMPO% = CODINT%(VDEVU$)
          CACONSU = 0
          TTYY$ = Space$(128)
          Call REPLA(TTYY$, VDEVU$, 1, 16)
          Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
          Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
          On Error Resume Next
          USOUNI = CACONSU / XVALO(Text1)
          On Error GoTo 0
          Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
          List1.AddItem TTYY$, ILISTA&
          List1.ListIndex = ILISTA&
          Text4 = ""
          Text4.SetFocus
       End If
    End If
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    ' validaciones
    '
    ' VALIDAR FECHA
    '\\\OT-06-0476-RG-04/12/06///
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% = 0 Or FECHA% > HOII% Then
       Call MENSERR(24, "Fecha No Válida")
       GoTo 99
    End If
    '\\\OT-06-0476-RG-FIN///
    '
    DEPO% = XVALO(Text2)
    If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Entrada")
       GoTo 99
    End If
    '
    DP2% = XVALO(Text3)
    If DP2% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DP2%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Consumo")
       GoTo 99
    End If
    '
    'VALIDAR CANTIDAD
    '\\\OT-06-0476-RG-04/12/06///
    VALONU = XVALO(Text1.TEXT)
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    '\\\OT-06-0476-RG-FIN///
    '
    CODICON$ = Label28
    IDMO$ = TRIM$(Label33)
    NORFA$ = Label32
    Call APERBASDAT("STOCKS")
    '
    If IDMO$ = "" Then
25     IDMO$ = NORFA$ + "/" + RANDSTRING$(-3)
       Movisto.FindFirst "IdenMovi = '" & IDMO$ & "'"
       If Movisto.NoMatch = False Then GoTo 25
    End If
    '
    CMOV$ = "CF"
    DOPRI$ = TRIM$(Label32)
    SQLT$ = "SELECT * FROM MOVISTO "
    SQLT$ = SQLT$ + " WHERE CodiMovi = '" & CMOV$ & "' "
    SQLT$ = SQLT$ + " AND DoPriCor = '" & DOPRI$ & "' "
    SQLT$ = SQLT$ + " AND IdenMovi = '" & IDMO$ & "' "
    Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenDynaset)
    '
    ' VERIFICAR NEGATIVOS
    'If Text8.Enabled = True Then     ' SOLO PRIMERA GRABACIÓN
       CAITMO% = 0
       For KKI% = 0 To List1.ListCount
          TTXX$ = List1.List(KKI%)
          CODE$ = TRIM$(Left$(TTXX$, 16))
          CIXI% = CODINT%(CODE$)
            CACONSACT = XVALO(Mid$(TTXX$, 57, 13))
            CACONSANT = 0   '
          CACONSU = CACONSACT - CANCONSANT ' PARA DESCONTAR EL CONSUMIDO ANTERIOR
          If CACONSU > 0.005 Then
             CAITMO% = CAITMO% + 1
             CIMOVIM%(CAITMO%) = CIXI%
             CAMOVIM(CAITMO) = (-1) * CACONSU
          End If
       Next KKI%
       '
       ' DESCONTAR DE CAMOVIM() LOS CONSUMOS CORRESPONDIENTES A LA MISMA o/f
       With MACONSUM
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             Do While Not .EOF
               CIXI1% = !Cod_Inte
               DPXX1% = !DepoMovi
               QTSAL1 = !Cantsalid
                 DEPMO% = DP2%: If DP2% = 99 Then DEPMO% = DEPOPRE%(CIXI1%)
               If DPXX1% = DEPMO% Then      ' OJO CON EL PREDETERMINADO
                 For JJ11% = 1 To CAITMO%
                   If CIMOVIM%(JJ11%) = CIXI1% Then
                     CAMOVIM(JJ11%) = CAMOVIM(JJ11%) + QTSAL1
                     Exit For
                   End If
                 Next JJ11%
               End If
             .MoveNext
             Loop
          End If
          On Error GoTo 0
       End With
       '
       If PUEDEMOVI%(DP2%) < 0 Then
         Call COMUNI("Transacción Cancelada por Control de Negativos.")
         GoTo 99
       End If
    'End If
    '
    ' GRABAR MOVIMIENTO DE ENTRADA
    CICON% = CODINT%(CODICON$)
    CODCON$ = CODEXT$(CICON%)
    ILOTE$ = Label32: If TRIM$(Text8) <> "" Then ILOTE$ = TRIM$(Text8)
    PDES$ = Text5
    PHAS$ = Text6
    CMOV$ = "OF"
    DOPRI$ = Label32
    DOSEC$ = "IF-" + TRIM$(Mid$(DOPRI$, 4))
    REFE$ = "Fabricación Propia " + CODICON$
    VUNI# = 0
    MONEX% = 1
    CANPRO = XVALO(Text1)
    GoSub GRAPRODUC
    '
    ' GRABAR CONSUMO
    CMOV$ = "CF"
    DOPRI$ = TRIM$(Label32)
    DOSEC$ = "PF-" + TRIM$(Mid$(Label32, 4))
    REFE$ = "CONSUMO " + CODICON$
    NORIT% = 0
    '
    SQLT$ = "SELECT * FROM MOVISTO "
    SQLT$ = SQLT$ + " WHERE CodiMovi = '" & CMOV$ & "' "
    SQLT$ = SQLT$ + " AND DoPriCor = '" & DOPRI$ & "' "
    SQLT$ = SQLT$ + " AND IdenMovi = '" & IDMO$ & "' "
    Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenDynaset)
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       CODE$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(CODE$)
       CACONSU = XVALO(Mid$(TTXX$, 57, 13))
       '
       SUMATRA = 0
       For KKL& = 129 To Len(TTXX$) Step 64
          TTYY$ = Mid$(TTXX$, KKL&, 64)
          LOTE$ = TRIM$(Left$(TTYY$, 16))
          If LOTE$ <> "" Then
             CANTIX = XVALO(TRIM$(Mid$(TTYY$, 17, 16)))
             If CANTIX >= 0.05 Then
                PCOM$ = TRIM$(Mid$(TTYY$, 33, 16))
                PFIN$ = TRIM$(Mid$(TTYY$, 49, 16))
                SUMATRA = SUMATRA + CANTIX
                NORIT% = NORIT% + 1
                GoSub GRACONSUM
             End If
          End If
       Next KKL&
       '
       CANTIX = CACONSU - SUMATRA
       If CANTIX >= 0.05 Then
          LOTE$ = ""
          PCOM$ = ""
          PFIN$ = ""
          NORIT% = NORIT% + 1
          GoSub GRACONSUM
       End If
       '
    Next KKI%
    '
    With MACONSUM
35    .FindFirst "NuOrItem > " & NORIT%
      If .NoMatch = False Then
         CIXI% = !Cod_Inte
         LOTE$ = !IdenLote
         .Delete
         Call ACSALOTE(CIXI%, LOTE$)
         GoTo 35
      End If
    End With

    ' ACTUALIZAR PENDIENTE DE O/F
    ACUPRODU = 0
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF' "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CICON% & " "
    Set ORYACIE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With ORYACIE
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            ACUPRODU = ACUPRODU + !CantIngre
         .MoveNext
         Loop
      End If
    End With
    '
    CIERRA% = 0
    If ACUPRODU >= XVALO(Label34) - 0.05 Then
        CIERRA% = 1
      Else
        MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
        If MDXX% = 2 Then CIERRA% = 1
    End If
    '
    Call APERBASDAT("ORFAB")
    With OrdeFabr
       .FindFirst "IdSubOr = '" & NORFA$ & "'"
       If .NoMatch = False Then
         .Edit
         !CanApro = ACUPRODU
         !CanRech = !CanRech + SCRAP
         If CIERRA% = 1 Then
             !Statuorf = 3         ' MARCA COMO CERRADA
           Else
             If !Statuorf > 2 Then !Statuorf = 1
         End If
         !FechCierr = CVDAT(FECHANUM(Label24))
         .Update
       End If
    End With
    '
    '25/04/07 (OT-07-0161)'PARA GUARDAR EL DEPOSITO DE ENTRADA Y EL DEPOSITO
    If CICON% > 0 Then    'DE CONSUMO EN EL MAESTRO DE ARTICULOS
        Master.FindFirst ("CODINT = " & CICON%)
        If Master.NoMatch = False Then
           Master.Edit
           Master!DepoEntra = XVALO(Text2)
           Master!DepoCon = XVALO(Text3)
           Master.Update
        End If
    End If
    '25/04/07 (FIN)
    '
    Call Label32_Change
    GoTo 99
    '
    '
GRAPRODUC:
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = 'OF') "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND IdenMovi = '" & IDMO$ & "'"
    Set PRODORFA = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With PRODORFA
       On Error Resume Next
       .MoveFirst
       If Err Then
            On Error GoTo 0
            .AddNew
          Else
            On Error GoTo 0
            .Edit
       End If
      !fechMov = CVDAT(FECHA%)
      !Cod_Inte = CICON%
      !COD_EXTE = CODCON$
      !Descrip = DESCRIT0$(CICON%)
      !IdenLote = TRIM$(ILOTE$)
      !PNL_Des = PDES$
      !PLC_Has = PHAS$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
      !NuOrItem = 0
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = VUNI#
      !MoneVal = MONEX%
      !TipoCam = TICAMONE(MONEX%)
         DEPOI% = DEPO%: If DEPO% = 99 Then DEPOI% = DEPOPRE%(CICON%)
      !DepoMovi = DEPOI%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = CANPRO
      !Cantsalid = 0
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
    End With
    '
    Call CREALOTIN(FECHA%, CICON%, ILOTE$, CDbl(CANPRO))
    '
Return
'
'
GRACONSUM:
    With MACONSUM
      LOTEA$ = ""
      .FindFirst "NuOrItem = " & NORIT%
      If .NoMatch = False Then
          LOTEA$ = !IdenLote
          CACONANT = !Cantsalid
          .Edit
        Else
          CACONANT = 0
          .AddNew
      End If
      !fechMov = CVDAT(FECHA%)
      !Cod_Inte = CIXI%
      !COD_EXTE = CODE$
      !Descrip = DESCRIT0$(CIXI%)
      !IdenLote = TRIM$(LOTE$)
      !PNL_Des = PCOM$
      !PLC_Has = PFIN$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
         LOTECON$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(CANTIX, "F8.1")
      !LotConsum = LOTECON$
      !NuOrItem = NORIT%
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = 0
      !MoneVal = 1
      !TipoCam = TICAMONE(1)
         DEPOI% = DP2%: If DP2% = 99 Then DEPOI% = DEPOPRE%(CICON%)
      !DepoMovi = DP2%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = 0
      !Cantsalid = Abs(CANTIX)
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
      '
      Call ACSALOTE(CIXI%, LOTE$)
      If LOTEA$ <> "" Then
        If LOTEA$ <> LOTE$ Then
          Call ACSALOTE(CIXI%, LOTEA$)
        End If
      End If
      '
      ' AQUI CANCELAR RESERVAS
    End With
Return
'
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
End Sub

Private Sub Command14_Click()
   '\\\OT-06-0476-RG-07/12/06///
  OPTRAZ04.Show 1
   '\\\OT-06-0476-RG-FIN///
End Sub



Private Sub Command16_Click()
   Command16.Enabled = False
   If TRIM$(Label32) = "" Then GoTo 99 ' DEBE HABER UNA o/f ACTIVA
   '
   FORIPRE$ = TRIM$(CONTROL$("TRAZABI", "FORTRAZA"))
   If FORIPRE$ <> "" Then
     If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
       GoTo 10
     End If
   End If
   Call MENSERR(24, "Formulario de Impresión Inexistente\o no Definido.")
   GoTo 99
   '
10 Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = Label32
   TIPODOC$ = "Registro de Trazabilidad"
   '
   CODPARAM$(1) = "REF-MAT1"
   CODPARAM$(2) = "MEDIDA-1"
   CODPARAM$(3) = "MEDIDA-2"
   CODPARAM$(4) = "MEDIDA-3"
   CAPARDOC% = 8
   '
   DOCPARAM$(1) = Text8
      If TRIM$(Text8) = "" Then
         DOCPARAM$(1) = TRIM$(Label32)
      End If
   DOCPARAM$(2) = TRIM$(Label30)
   DOCPARAM$(3) = TRIM$(Text5)
   DOCPARAM$(4) = TRIM$(Text6)
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "N-PLANO"
   CODTABU1$(4) = "R-PLANO"
   CODTABU1$(5) = "F-PLANO"
   CODTABU1$(6) = "RASO-CLI"
   CODTABU1$(7) = "Z-PLANO"
   CODTABU1$(8) = "FECH-REC"
   CODTABU1$(9) = "FECH-VEN"
   CODTABU1$(10) = "NRO-LOTE"
   CODTABU1$(11) = "N-AUTENT"
   CODTABU1$(12) = "REF-MAT2"
   CODTABU1$(13) = "N-HPROF"
   CODTABU1$(14) = "N-HINSP"
   CACOLTAB1% = 14
   '
   CAITAB1% = 0
   For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      CODE$ = TRIM$(Left$(Z$, 16))
      CIX0% = CODINT%(CODE$)
      '
      TTXX$ = Z$
      For KKL& = 129 To Len(TTXX$) Step 64
         TTYY$ = Mid$(TTXX$, KKL&, 64)
         IDLOTE$ = TRIM$(Left$(TTYY$, 16))
         If IDLOTE$ <> "" Then
           CACOLO = XVALO(TRIM$(Mid$(TTYY$, 17, 16)))
           If CACOLO >= 0.05 Then
             PDES$ = TRIM$(Mid$(TTYY$, 33, 16))
             PHAS$ = TRIM$(Mid$(TTYY$, 49, 16))
             '
             TIPOX$ = ""
             MARCAX$ = ""
             PARTNUX$ = ""
             PROVX$ = ""
             FINGRE = 0
             FEFAB = 0
             TRACLIX$ = ""
             VERIX$ = ""
             SQLX$ = "SELECT * FROM BOLRECEP"
             SQLX$ = SQLX$ + " WHERE IdenLote = '" & IDLOTE$ & "'"
             Set BRSELE = Stocks.OpenRecordset(SQLX, dbOpenSnapshot)
             With BRSELE
               On Error Resume Next
               .MoveFirst
               If Err Then
                  On Error GoTo 0
                  ' Suprime mensaje de error porque puede provenir de una O/F
                  ' Call MENSERR(24, "No Encontrado Lote '" + IDLOTE$ + "'")
                  GoTo 30
               End If
               On Error GoTo 0
               '
               TIPOX$ = "": MARCAX$ = "": PARTNUX$ = "": PROVX$ = "": TRACLIX$ = "": VERIX$ = "": COLOTX$ = ""
               On Error Resume Next
               TIPOX$ = TRIM$(!TIPO)
               MARCAX$ = !MARCA
               PARTNUX$ = !NUMEPARTE
               PROVX$ = !RASO_PROV
               FINGRE = CVINT(!FechRecep)
               FEFAB = CVINT(!FECHFAB)
               COLOTX$ = !COD_CLI
               TRACLIX$ = !TRAZ_CLI
               VERIX$ = !USUAPRO
               On Error GoTo 0
               '
30             CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODE$
               TETABU1$(2, CAITAB1%) = DESCRIT0$(CIX0%)
                 If TIPOX$ = "" Then TIPOX$ = ATRIB1$(CIX0%)
               TETABU1$(3, CAITAB1%) = TIPOX$
               TETABU1$(4, CAITAB1%) = MARCAX$
               TETABU1$(5, CAITAB1%) = PARTNUX$
               TETABU1$(6, CAITAB1%) = PROVX$
               If Left$(IDLOTE$, 2) = "BR" Then
                  TETABU1$(7, CAITAB1%) = IDLOTE$
               End If
               TETABU1$(8, CAITAB1%) = FECHATEX$(FINGRE)
               TETABU1$(9, CAITAB1%) = FECHATEX$(FEFAB)
               If Left$(IDLOTE$, 2) <> "BR" Then
                    TETABU1$(10, CAITAB1%) = IDLOTE$              ' NUMERO DE LOTE
                  Else
                    TETABU1$(10, CAITAB1%) = COLOTX$
               End If
               TETABU1$(11, CAITAB1%) = TRACLIX$
               TETABU1$(12, CAITAB1%) = VERIX$
               TETABU1$(13, CAITAB1%) = PDES$
               TETABU1$(14, CAITAB1%) = PHAS$
             End With
           End If
         End If
      Next KKL&
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   '
99 Screen.MousePointer = 1
   Command16.Enabled = True
   '
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   SUMATRA = 0
   For KKU& = 0 To 5
      TTYY$ = Space$(64)
      If TRIM$(Label23(KKU&)) <> "" Then
        If XVALO(Text11(KKU&)) >= 0.05 Then
          SUMATRA = SUMATRA + XVALO(Text11(KKU&))
        End If
      End If
   Next KKU&
   If SUMATRA > XVALO(Text4) + 0.005 Then
      Call MENSERR(24, "Imposible Registrar.\Los Lotes Trazados Superan el Total Consumido.")
      GoTo 99
   End If
   '
   TTXX$ = Left$(List1.TEXT + Space$(128), 128)
   CACONSU = XVALO(Text4)
   Call REPLA(TTXX$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
   On Error Resume Next
   USOUNI = CACONSU / XVALO(Text1)
   On Error GoTo 0
   Call REPLA(TTXX$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
   Call REPLA(TTXX$, FORMATNUM$(SUMATRA, "F12.1"), 71, 12)
   '
   For KKU& = 0 To 5
      TTYY$ = Space$(64)
      If TRIM$(Label23(KKU&)) <> "" Then
        If XVALO(Text11(KKU&)) >= 0.05 Then
          Call REPLA(TTYY$, TRIM$(Label23(KKU&)), 1, 16)
          Call REPLA(TTYY$, Text11(KKU&), 17, 16)
          Call REPLA(TTYY$, Text16(KKU&), 33, 16)
          Call REPLA(TTYY$, Text17(KKU&), 49, 16)
          TTXX$ = TTXX$ + TTYY$
        End If
      End If
   Next KKU&
   '
   ILISTA& = XVALO(Label8)
   If List1.ListCount > 0 Then
      If ILISTA& >= 0 Then
        List1.RemoveItem (ILISTA&)
        List1.AddItem TTXX$, ILISTA&
        List1.ListIndex = ILISTA&
      End If
   End If
   '
99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call CommSel("SELFECHA", Command4, Label24)
End Sub

Private Sub Command5_Click()
    '
    Call CommSel("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)", Command5, Label32)
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    '
End Sub

Sub TAMARCAS()
    '
    CI% = CODINT%(Label27)
    If CI% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      CI% = CODINT(COD1$)
20    If CI% > 0 Then
         Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CI%))
         VTB% = VENTABU%("TAMARCAS")
         If VTB% >= 0 Then
            Call FILTERUPDATE("TAMARCAS", "!TAMARCAS", 4, MKI$(CI%))
            If CODINT%(Label27) < 1 Then GoTo 10
         End If
      End If
    End If
    '
End Sub
'
Private Sub Command6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command6_Click(Index)
   End If
End Sub

Private Sub Command7_Click()
    Call CommSel("TADEPOSI", Command7, Text2)
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call TRAZABI04.Command7_Click
   Command8.Enabled = True
End Sub

'
Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text3)
End Sub

Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function

Private Sub Command6_Click(Index As Integer)
    If TRIM$(Label27) = "" Or CODINT%(Label27) < 0 Then
       Exit Sub
    End If
    '
    For KKU& = 0 To Index - 1
       If Label23(KKU&) = "" Then
          Call MENSERR(24, "No Dejar Campos no Informados")
          Exit Sub
       End If
    Next KKU&
    '
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    If VDEVU$ <> "" Then
      For KKU& = 0 To 5
        If KKU& <> Index Then
          If Label23(KKU&) = VDEVU$ Then
             Call MENSERR(24, "No Válido Repetir Lotes")
             Exit Sub
          End If
        End If
      Next KKU&
      Label23(Index) = VDEVU$
    End If
    '
    CAMAXI = XVALO(Text4)
    For KKU& = 0 To 5
       If KKU& <> Index Then
          CAMAXI = CAMAXI - XVALO(Text11(KKU&))
       End If
    Next KKU&
    If CAMAXI < 0 Then CAMAXI = 0
    '
    If XVALO(Mid$(VDEV2$, 21, 12)) < CAMAXI Then
       CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
    End If
    If TRIM$(Text11(Index)) = "" Then
       Text11(Index) = FORMATNUM$(CAMAXI, "F10.1")
    End If
    '
    Text11(Index).SetFocus
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       Frame9.Visible = True
    End If
    '
    Call FRESHALL
    '
    If Left$(APLINVO$, 7) = "OFABRIC" Then
       Label32 = TRIM$(Mid$(APLINVO$, 8))
       If Label32 <> "" Then
          If ECOARCH$("ECORDEP", Label32) <> "" Then
            HOII% = HOY(HOS$)
            Label24 = HOS$
          End If
       End If
    End If
    '
End Sub

Private Sub Label27_Change()
   CIXI% = CODINT%(Label27)
   If CIXI% > 0 Then
      'Call CREALOTES(CIXI%)
      'Call VERSALIDA(CIXI%)
      Call ARMALILOT(CIXI%)
      Screen.MousePointer = 1
   End If
End Sub

Sub ARMALILOT(CIXI%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + "ORDER BY FechAlta desc "
    Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LotIngTemp
25     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TTXX$ = REBLA$
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, TRIM$(!IdenLote), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F12,1"), 25, 12)
           Call REPLA(TTXX$, FORMATNUM$(!CANSALDO, "F12,1"), 37, 12)
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
         .MoveNext
         Loop
       End If
    End With
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub
'
Private Sub Label27_Click()

End Sub

Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
    List2.Clear
    Label17 = ""
    Label20 = ""
    Label21 = ""
    Label33 = ""
    HOII% = HOY(HOS$)
    Label24 = HOS$
    Text8.Enabled = True
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    List1.Clear
    List2.Clear
    Label28 = ""
    Label30 = ""
    Text8 = NORFA$
    Text1 = ""
    Text5 = ""
    Text6 = ""
    Text2 = ""
    Text3 = ""
    Label34 = ""
    Label27 = ""
    Label6 = ""
    Label16 = ""
    Text4 = ""
    '
    For KKU& = 0 To 5
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Text16(KKU&) = ""
       Text17(KKU&) = ""
    Next KKU&
    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
         Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With ORDEFABTMP
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "Orden de Fabricación no Válida. Consulte.")
              Call FINAL("")
           End If
           On Error GoTo 0
           CIXI% = !Cod_Inte
           Label28 = CODEXT$(CIXI%)
           Label30 = DESCRIT0$(CIXI%)
              CAPENDE# = !CanProy - !CanApro: If CAPENDE < 0 Then CAPENDE = 0
              If !Statuorf >= 3 Then CAPENDE# = 0
           Text1 = FORMATNUM$(CAPENDE#, "F8.1")
           Label34 = FORMATNUM$(!CanProy, "F8.1")
           '
           '25/04/07 (OT-07-0161)
           Text2 = Str(DEPENART%(CIXI%)) 'PARA MOSTRAR EL DEPOSITO DE ENTRADA Y DE CONSUMO
           Text3 = Str(DEPCONART%(CIXI%)) 'ASIGNADO AL ARTICULO
           '25/04/07 (FIN)
           '
           Call CARLICONSU(1)
         End With
         '
         Call APERBASDAT("STOCKS")
         SQLX$ = "SELECT * FROM MOVISTO "
         SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF' "
         SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
         SQLX$ = SQLX$ + " AND Cod_Inte = " & CIXI% & " "
         SQLX$ = SQLX$ + " ORDER BY FechMov ASC , FeReMovi ASC "
         Set ORYACIE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
         '
         With ORYACIE
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              On Error GoTo 0
              Do While Not .EOF
                 ' verificar si tiene IdenMovi
                 If IsNull(!IdenMovi) = True Or !IdenMovi = "" Then
25                  IDMO$ = NORFA$ + "/" + RANDSTRING$(-3)
                    Movisto.FindFirst "IdenMovi = '" & IDMO$ & "'"
                    If Movisto.NoMatch = False Then GoTo 25
                    .Edit
                    !IdenMovi = IDMO$
                    .Update
                 End If
                 '
                 On Error Resume Next
                 IDMO$ = !IdenMovi
                 PDES$ = !PNL_Des
                 PHAS$ = !PLC_Has
                 On Error GoTo 0
                 TTXX$ = Space$(256)
                 Call REPLA(TTXX$, FECHATEX$(CVINT(!fechMov)), 1, 8)
                 Call REPLA(TTXX$, FORMATNUM$(!CantIngre, "F9.1"), 10, 9)
                 DIHORE$ = Format$(!FeReMovi, "dd/mm/yy  hh:mm:ss")
                 DIRE$ = Left$(DIHORE$, 8)
                 HORE$ = Mid$(DIHORE$, 11, 8)
                 NUSE& = 10000 * XVALO(Left$(HORE$, 2)) + 100 * XVALO(Mid$(HORE$, 4, 2)) + XVALO(Mid$(HORE$, 7, 2))
                 Call REPLA(TTXX$, DIRE$, 33, 8)
                 Call REPLA(TTXX$, Str$(NUSE&), 43, 10)
                 Call REPLA(TTXX$, IDMO$, 61, 20)
                 Call REPLA(TTXX$, !IdenLote, 81, 32)
                 Call REPLA(TTXX$, TRIM$(Str$(!DepoMovi)), 113, 4)
                 Call REPLA(TTXX$, PDES$, 117, 32)
                 Call REPLA(TTXX$, PHAS$, 149, 32)
                 List2.AddItem TTXX$
              .MoveNext
              Loop
           End If
           On Error GoTo 0
         End With
    End If
    Screen.MousePointer = 1
End Sub

Sub CARLICONSU(MODO%)
    ' MODO%=1 DESDE LAS RESERVAS
    List1.Clear
    Label8 = (-1)
    CACONJU# = XVALO(Text1)
    If CACONJU# < 0.1 Then Exit Sub
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & Label32 & "'"
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    CAIT% = 0
    With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While (.EOF = False)
            CAIT% = CAIT% + 1
            CIJ%(CAIT%) = !Cod_Inte
            USOI(CAIT%) = !UsoUnit
          .MoveNext
          Loop
       End If
       On Error GoTo 0
    End With
    '
    If CAIT% < 1 Then    ' No Existen Reservas efectuadas.
       ' arma composicion segun formula standard
       CODPRO$ = Label28
       Call LEEEXPLOMUL(CODPRO$, 1, 1)
    End If
    '
    If CAIT% < 1 Then             ' producto sin estructura
      Call MENSERR(24, "No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
    End If
    '
    FIN& = CAIT%
    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
          If Abs(USOI(I1%)) > 0 Then
            ZZ$ = Space$(128)
            Call REPLA(ZZ$, CODEXT$(CIJ%(I1%)), 1, 16)
            Call REPLA(ZZ$, DESCRIT$(CIJ%(I1%)), 17, 40)
              NEBRU = CACONJU# * USOI(I1%)
            Call REPLA(ZZ$, FORMATNUM$(NEBRU, "F13.1"), 57, 13)
            Call REPLA(ZZ$, FORMATNUM$(USOI(I1%), "F16.6"), 97, 16)
            List1.AddItem ZZ$
          End If
       End If
    Next I1%
    Screen.MousePointer = 1
    '
End Sub

Private Sub Label32_DblClick()
    ' USAR PARA LIMPIAR TODO
    NOFA$ = Label32
    Label32 = ""
    Label32 = NOFA$
End Sub

'
Private Sub List1_Click()
    
    TTXX$ = List1.TEXT
    COD$ = Mid(TTXX$, 1, 16)
    CI% = CODINT%(COD$)
    DESC$ = DESCRIT$(CI%)
    CantNom$ = Mid(TTXX$, 57, 13)
    Label27.Caption = COD$
    Label6.Caption = DESC$
    Text4 = TRIM$(CantNom$)
    Label8 = List1.ListIndex
    '
    For KKU& = 0 To 5
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Text16(KKU&) = ""
       Text17(KKU&) = ""
    Next KKU&
    '
    KKU& = 0
    For KKL& = 129 To Len(TTXX$) Step 64
       TTYY$ = Mid$(TTXX$, KKL&, 64)
       Label23(KKU&) = RTrim(Left$(TTYY$, 16))
       Text11(KKU&) = TRIM$(Mid$(TTYY$, 17, 16))
       Text16(KKU&) = TRIM$(Mid$(TTYY$, 33, 16))
       Text17(KKU&) = TRIM$(Mid$(TTYY$, 49, 16))
       KKU& = KKU& + 1
    Next KKL&
    '
End Sub

Private Sub List2_Click()
    Call List2_DblClick
End Sub

Private Sub List2_DblClick()
    TTXX$ = List2.TEXT
    NORFA$ = Label32
    Label24 = Left$(TTXX$, 8)
    FEMOVI = CDbl(CVDAT(FECHANUM(Label24)))
    Text1 = TRIM$(Mid$(TTXX$, 10, 9))
    If XVALO(Text1) < 0.1 Then Text1 = "1.0"
    NUSe0& = XVALO(TRIM$(Mid$(TTXX$, 43, 10)))
    '
    Label21 = Left$(TTXX$, 8)
    Label17 = Mid$(TTXX$, 33, 8)
    Label20 = Mid$(TTXX$, 43, 10)
    Label33 = TRIM$(Mid$(TTXX$, 61, 20))
    Text8 = TRIM$(Mid$(TTXX$, 81, 32))
    Text2 = TRIM$(Mid$(TTXX$, 113, 4))
    Text5 = TRIM$(Mid$(TTXX$, 117, 32))
    Text6 = TRIM$(Mid$(TTXX$, 149, 32))
    Text8.Enabled = False
    '
    If Label33 = "" Then             ' no tiene IdenMovi
       Call MENSERR(24, "Imposible Determinar Identificador de Movimiento - Consulte !")
       Call FINAL("")
    End If
    '
    List1.Clear
    Label27.Caption = ""
    Label6.Caption = ""
    Text4 = ""
    Label8 = (-1)
    Label16 = ""
    '
    For KKU& = 0 To 5
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Text16(KKU&) = ""
       Text17(KKU&) = ""
    Next KKU&
    '
    IDMO$ = Label33
    Call APERBASDAT("STOCKS")
    SQLT$ = "SELECT * FROM MOVISTO "
    SQLT$ = SQLT$ + " WHERE CodiMovi = 'CF' "
    SQLT$ = SQLT$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLT$ = SQLT$ + " AND IdenMovi = '" & IDMO$ & "' "
    SQLT$ = SQLT$ + "ORDER BY NuOrItem ASC "
    Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenSnapshot)
    '
    On Error Resume Next
    MACONSUM.MoveFirst
    If Err > 0 Then
       On Error GoTo 0
       GoSub RECUIDENMO
       Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenSnapshot)
    End If
    On Error GoTo 0
    '
15  With MACONSUM
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
             CICOMPO% = !Cod_Inte
             CODECO$ = TRIM$(CODEXT$(CICOMPO%))
             CACONSU = !Cantsalid
             DepoCon% = !DepoMovi
             '
             For JJ& = 0 To List1.ListCount - 1
               If TRIM$(Left$(List1.List(JJ&), 16)) = CODECO$ Then
                 TTYY$ = List1.List(JJ&)
                 RELI& = JJ&
                 CACONSU = CACONSU + XVALO(Mid$(TTYY$, 57, 13))
                 Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
                 USOUNI = CACONSU / XVALO(Text1)
                 Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
                 List1.RemoveItem RELI&
                 List1.AddItem TTYY$, RELI&
                 GoTo 20
               End If
             Next JJ&
             RELI& = List1.ListCount
             TTYY$ = Space$(128)
             Call REPLA(TTYY$, CODECO$, 1, 16)
             Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
             Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
             USOUNI = CACONSU / XVALO(Text1)
             Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
             List1.AddItem TTYY$
             '
20           SUMATRA = XVALO(Mid$(TTYY$, 71, 12))
             If IsNull(!LotConsum) = False Then
                If TRIM$(!LotConsum) <> "" Then
                   For KKL& = 1 To Len(!LotConsum) Step 32
                     TTZZ$ = Space$(64)
                     IDLOCO$ = TRIM$(Mid$(!LotConsum, KKL&, 24))
                     CACOLO = XVALO(Mid$(!LotConsum, KKL& + 24, 8))
                     If IDLOCO$ <> "" Then SUMATRA = SUMATRA + CACOLO
                     Call REPLA(TTZZ$, IDLOCO$, 1, 16)
                     Call REPLA(TTZZ$, FORMATNUM$(CACOLO, "F12.1"), 17, 16)
                     On Error Resume Next
                     Call REPLA(TTZZ$, !PNL_Des, 33, 16)
                     Call REPLA(TTZZ$, !PLC_Has, 49, 16)
                     On Error GoTo 0
                     TTYY$ = TTYY$ + TTZZ$
                   Next KKL&
                   '
                   Call REPLA(TTYY$, FORMATNUM$(SUMATRA, "F12.1"), 71, 12)
                   List1.RemoveItem RELI&
                   List1.AddItem TTYY$, RELI&
                End If
             End If
          .MoveNext
          Loop
       End If
    End With
    '
    Text3 = TRIM$(Str$(DepoCon%))
    '
    Screen.MousePointer = 1
    Exit Sub
    '
RECUIDENMO:
    ' RESTAURA IDENTIFICADOR DE MOVIMIENTO
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi = 'CF' "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND FechMov = " & FEMOVI & " "
    Set MACONSUM = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MACONSUM
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
             DIHORE$ = Format$(!FeReMovi, "dd/mm/yy  hh:mm:ss")
             DIRE$ = Left$(DIHORE$, 8)
             HORE$ = Mid$(DIHORE$, 11, 8)
             NUSE& = 10000 * XVALO(Left$(HORE$, 2)) + 100 * XVALO(Mid$(HORE$, 4, 2)) + XVALO(Mid$(HORE$, 7, 2))
             If DIRE$ = Label17 Then
               If Abs(NUSE& - NUSe0&) < 5 Then
                 If IsNull(!IdenMovi) = False Then
                    If TRIM$(!IdenMovi) <> "" Then
                      GoTo 39
                    End If
                 End If
                 .Edit
                 !IdenMovi = IDMO$
                 .Update
               End If
             End If
39        .MoveNext
          Loop
       End If
    End With
Return
'
End Sub

Sub LIMPLAB()

    Label27.Caption = ""
    Label6.Caption = ""
    Text4.TEXT = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text12 = ""
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Call Text1_LostFocus
      Text5.SetFocus
   End If
End Sub

Private Sub Text1_LostFocus()
   If XVALO(Text1) < 0.1 Then
      Call MENSERR(24, "Mandatorio Informar Cantidad")
      Text1 = "1.0"
   End If
   Text1 = FORMATNUM(Text1, "F8.1")
   '
   For JJ& = 0 To List1.ListCount - 1
      TTYY$ = List1.List(JJ&)
      RELI& = JJ&
      CACONSU = XVALO(Text1) * XVALO(Mid$(TTYY$, 97, 16))
      Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
      List1.RemoveItem RELI&
      List1.AddItem TTYY$, RELI&
   Next JJ&
   List1.Refresh
   '
   Label27 = ""
   Label6 = ""
   Text4 = ""
   Label8 = (-1)
   '
End Sub

Private Sub Text11_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 5
       If Label23(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text11(KKI%))
         Else
           Text11(KKI%) = ""
       End If
    Next KKI%
    Label16 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
End Sub

Private Sub Text16_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text17(Index).SetFocus
    End If
End Sub

Private Sub Text17_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 5 Then
          Command6(Index + 1).SetFocus
       End If
    End If
End Sub

Private Sub Text2_Change()
  '\\\OT-06-0476-RG-04/12/06///
  DEPO% = XVALO(Text2)
  Label9.Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
  '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Text2_DblClick()
  '\\\OT-06-0476-RG-04/12/06///
  Call Command7_Click
  '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Private Sub Text3_Change()
  '\\\OT-06-0476-RG-04/12/06///
  DEPO% = XVALO(Text3)
  Label11.Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
  '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Text3_DblClick()
  '\\\OT-06-0476-RG-04/12/06///
  Call Command9_Click
  '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

    If KeyAscii% = 13 Then
       List1.SetFocus
    End If
    
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

    If KeyAscii% = 13 Then
       Call Text4_LostFocus
    End If
    
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      Text8.SetFocus
   End If
   
End Sub

Private Sub Text7_lostfocus()
  '
  Static MODORA%
  If MODORA% < 1 Then
     MODORA% = 1
     If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
        MODORA% = 2
     End If
  End If
  '
  If XVALO(Text7.TEXT) > 0 Then
    Text7.TEXT = REPLACAR$(Text7.TEXT, ".", ":")
    If InStr(1, Text7.TEXT, ".") = 0 Then
      If InStr(1, Text7.TEXT, ":") = 0 Then
        Text7.TEXT = Text7.TEXT + ":0"
      End If
    End If
    '
    If MODORA% = 2 Then
       TTXX1$ = FORMATNUM$(REPLACAR$(Text7.TEXT, ":", "."), "F7.2")
       Text7 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
       Exit Sub
    End If
    '
    On Error Resume Next
      D1 = CDate(Text7.TEXT)
    On Error GoTo 0
    Text7.TEXT = Format$(D1, "hh:mm")
    '
  End If
  '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text2.SetFocus
    End If

End Sub

Private Sub Text11_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text11_LostFocus(Index)
       Text16(Index).SetFocus
    End If
End Sub

Private Sub Text11_LostFocus(Index As Integer)
    Text11(Index) = TRIM$(FORMATNUM$(Text11(Index), "F10.1"))
End Sub

Private Sub Text4_LostFocus()
   '
   Text4 = TRIM$(FORMATNUM$(Text4, "F12.1"))
   '
   If TRIM$(Label27) <> "" Then
      If XVALO(Text4) < 0.05 Then
         For JJ& = 0 To List1.ListCount - 1
            TTYY$ = List1.List(JJ&)
            If TRIM$(Left$(TTYY$, 16)) = TRIM$(Label27) Then
               List1.RemoveItem JJ&
               GoTo 80
            End If
         Next JJ&
         '
80       Label27 = ""
         Label6 = ""
         Text4 = ""
         Label8 = (-1)
      End If
   End If
   '
   List1.Refresh
   '
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub
