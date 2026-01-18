VERSION 5.00
Begin VB.Form TRAZORF04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   7600
      Left            =   9765
      ScaleHeight     =   7545
      ScaleWidth      =   1035
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   7670
         Left            =   0
         Picture         =   "TRAZORF04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cantidad"
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
      TabIndex        =   87
      Top             =   210
      Width           =   2010
      Begin VB.Label Label38 
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
         Left            =   945
         TabIndex        =   93
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label37 
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
         Left            =   945
         TabIndex        =   92
         Top             =   735
         Width           =   915
      End
      Begin VB.Label Label34 
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
         Left            =   945
         TabIndex        =   91
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   -525
         TabIndex        =   90
         Top             =   1095
         Width           =   1360
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyect:"
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
         Left            =   -525
         TabIndex        =   89
         Top             =   465
         Width           =   1360
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
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
         Left            =   -525
         TabIndex        =   88
         Top             =   780
         Width           =   1360
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame9"
      Height          =   540
      Left            =   2940
      TabIndex        =   80
      Top             =   4095
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label20 
         Height          =   225
         Left            =   4725
         TabIndex        =   85
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label17 
         Height          =   225
         Left            =   3570
         TabIndex        =   84
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label21 
         Height          =   225
         Left            =   2415
         TabIndex        =   83
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label33 
         Height          =   225
         Left            =   735
         TabIndex        =   82
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label Label8 
         Height          =   225
         Left            =   105
         TabIndex        =   81
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8295
      TabIndex        =   77
      Top             =   210
      Width           =   1275
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
         Left            =   315
         TabIndex        =   4
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
         Left            =   315
         TabIndex        =   3
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   315
         Width           =   1335
      End
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
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   46
      Top             =   1785
      Width           =   9360
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   8900
         TabIndex        =   76
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   7570
         TabIndex        =   75
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   6250
         TabIndex        =   71
         Top             =   450
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1700
         Left            =   1850
         TabIndex        =   70
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
         TabIndex        =   5
         Top             =   630
         Width           =   9000
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo                   Descripción                                                               Reserva            Qt.Trz."
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
         TabIndex        =   51
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
      TabIndex        =   39
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Listado Materiales Reservados por Lote"
         Top             =   2550
         Width           =   650
      End
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   4500
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3850
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
         Height          =   1005
         Left            =   105
         Picture         =   "TRAZORF04.frx":2DEE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Guardar Cierre de Orden de Fabricacion"
         Top             =   5565
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
         Height          =   660
         Left            =   105
         Picture         =   "TRAZORF04.frx":30F8
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3200
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
         Picture         =   "TRAZORF04.frx":3402
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZORF04.frx":370C
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Impresión de Registro de Trazabilidad"
         Top             =   1900
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   40
         Top             =   6720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   41
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
         Picture         =   "TRAZORF04.frx":3A16
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRAZORF04.frx":3B60
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7030
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   42
      Top             =   4200
      Width           =   9360
      Begin VB.CommandButton Command15 
         Caption         =   "Auto F.I.F.O."
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
         Left            =   210
         TabIndex        =   97
         Top             =   2415
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F.I.F.O."
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
         TabIndex        =   96
         Top             =   2070
         Width           =   855
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
         Height          =   750
         Left            =   8295
         Picture         =   "TRAZORF04.frx":5A82
         Style           =   1  'Graphical
         TabIndex        =   30
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   72
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
         TabIndex        =   26
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
         TabIndex        =   22
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
         TabIndex        =   18
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
         TabIndex        =   14
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
         TabIndex        =   10
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
         TabIndex        =   6
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   11
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   7
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   74
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Q.Trazada:"
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
         Left            =   6825
         TabIndex        =   73
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   1260
         Width           =   1965
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qt Requerida:"
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
         Left            =   3675
         TabIndex        =   57
         Top             =   495
         Width           =   1755
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   47
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
         TabIndex        =   43
         Top             =   1260
         Width           =   1965
      End
   End
   Begin VB.Label Label7 
      Height          =   225
      Left            =   210
      TabIndex        =   95
      Top             =   1470
      Visible         =   0   'False
      Width           =   960
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
      TabIndex        =   69
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
      TabIndex        =   60
      ToolTipText     =   "Doble-Click para Blanquear"
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
      TabIndex        =   59
      Top             =   1220
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
      TabIndex        =   58
      Top             =   1155
      Width           =   4635
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
      TabIndex        =   56
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
      TabIndex        =   55
      Top             =   800
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Lanzamiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3465
      TabIndex        =   50
      Top             =   165
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   45
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
      TabIndex        =   44
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu consylist 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu printregistroytrazabilidad 
         Caption         =   "Impresión de Registro de Trazabilidad"
      End
      Begin VB.Menu materialreservado 
         Caption         =   "Listado Materiales Reservados por Lote "
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu controlmovimientosdestockrepetidos 
         Caption         =   "Control de Movimientos de Stock Repetidos o Duplicados"
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu kardex 
         Caption         =   "Ficha Kardex"
      End
   End
End
Attribute VB_Name = "TRAZORF04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%

Private Sub Command1_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then
      If COMALTER%("¿ Realmente Cancela la Asignación\Automática de Lotes Propuesta ?\\No, Conservarla\Si, Cancelarla") = 2 Then GoTo 10
      Exit Sub
   End If
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command10_Click()

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
    '
    ' VALIDAR FECHA
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% < 1 Then
       Call MENSERR(24, "Fecha de Lanzamiento No Válida")
       GoTo 99
    End If
    '
    'VALIDAR CANTIDAD
    VALONU = XVALO(Label38)
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    '
    CODICON$ = Label28
    IDMO$ = TRIM$(Label33)
    NORFA$ = Label32
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("RESERVA")
    '
'    If IDMO$ = "" Then
'25     IDMO$ = NORFA$ + "/" + RANDSTRING$(-3)
'       Movisto.FindFirst "IdenMovi = '" & IDMO$ & "'"
'       If Movisto.NoMatch = False Then GoTo 25
'    End If
'    '
    ' GRABAR RESERVA
' CREAR UN RECORDSET MARESER PARA GRABAR LA RESERVA
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
    Set MARESER = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MARESER
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
             .Edit
             !Status = 9
             .Update
          .MoveNext
          Loop
       End If
       On Error GoTo 0
    End With
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       CODE$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(CODE$)
       CACONSU = XVALO(Mid$(TTXX$, 57, 13))
       USOUN = CACONSU / VALONU
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
                GoSub GRARESERVA
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
          GoSub GRARESERVA
       End If
       '
    Next KKI%
    '
24  SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
    SQLX$ = SQLX$ + "AND Status = 9 "
    Set MARESANT = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MARESANT
       On Error Resume Next
25     .MoveFirst
       If Err < 1 Then
          CIXI% = !Cod_Inte
          LTXX$ = !IdenLote
          .Delete
          Call ACSALOTE(CIXI%, LTXX$, "NOMESS")
          GoTo 24
       End If
       On Error GoTo 0
    End With
    '
    GoTo 99
    '
GRARESERVA:
    With MARESER
      .AddNew
      !Cod_Inte = CIXI%
      !COD_EXTE = CODEXT$(CIXI%)
      !DESCRIP = DESCRIT0$(CIXI%)
      !IdSubOr = NORFA$
      !Status = 0
      !NuOrItem = NORIT%
      !FechRes = CVDAT(FECHANUM(Label24))
      !CantRes = CANTIX
      !CantCons = 0
      !Saldo_Entre = CANTIX
        LOTEA$ = ""
        On Error Resume Next
        LOTEA$ = !IdenLote
        On Error GoTo 0
      !IdenLote = LOTE$
      !Deposito = DEPOPRE(CIXI%)
      !UsoUnit = USOUN
      !Observa = ""
      !PNL_DES = PCOM$
      !PLC_HAS = PFIN$
      .Update
      '
      Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
      '
    End With
Return
'
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then
       Call CIERRARCH("*.*")
       Hide
    End If
    '
End Sub

Sub Command14_Click()
    '\\\OT-06-0476-RG-07/12/06///
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    OPTRAZ04.Show 1
    '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    BAUTOTRA% = 1
    For KKU& = 1 To List1.ListCount
      List1.Selected(KKU& - 1) = True
      Call AUTOTRA
    Next KKU&
    BAUTOTRA% = 0
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
   '
   Command16.Enabled = False
   NORFA$ = TRIM$(Label32)
   If NORFA$ = "" Then
      Call ABREORFAB
      Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
      NORFA$ = TRIM$(VALACT1$("ECORDEF"))
   End If
   If NORFA$ = "" Then GoTo 99
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
   Dim CODEX$(4096), CACONSU(4096), IDLOT$(4096), PNLD$(4096), PLCH$(4096)
   '
10 Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   FEX = HOY(HOS$)
   FECHDOC$ = HOS$
   NUMEDOC$ = NORFA$
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
         DOCPARAM$(1) = NORFA$
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
   CODTABU1$(15) = "CANTIDAD"
   CACOLTAB1% = 15
   '
   Call APERBASDAT("STOCKS")
   SQLT$ = "SELECT * FROM MOVISTO "
   SQLT$ = SQLT$ + " WHERE CodiMovi = 'CF' "
   SQLT$ = SQLT$ + " AND DoPriCor = '" & NORFA$ & "' "
   SQLT$ = SQLT$ + "ORDER BY NuOrItem ASC "
   Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenSnapshot)
   '
15 JJ& = 0
   With MACONSUM
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          If !Cantsalid >= 0.05 Then
            JJ& = JJ& + 1
            CICOMPO% = !Cod_Inte
            CODEX$(JJ&) = CODEXT$(CICOMPO%)
            CACONSU(JJ&) = !Cantsalid
            IDLOT$(JJ&) = !IdenLote
              PNLD(JJ&) = "": PLCH(JJ&) = ""
              On Error Resume Next
            PNLD(JJ&) = !PNL_DES
            PLCH(JJ&) = !PLC_HAS
              On Error GoTo 0
          End If
        .MoveNext
        Loop
     End If
   End With
   '
   ' AQUI AGREGAR LOS DE RESERVAS
   Call APERBASDAT("RESERVA")
   SQLX$ = " SELECT * FROM RESERVA "
   SQLX$ = SQLX$ + "WHERE IdSubOr = '" & NORFA$ & "' "
   Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While (.EOF = False)
            CACOLO = !CantRes - !CantCons: If CACOLO < 0 Then CACOLO = 0
            If CACOLO >= 0.05 Then
               If TRIM$(!IdenLote) <> "" Then    ' van solo las reservas trazadas
                  JJ& = JJ& + 1
                  CICOMPO% = !Cod_Inte
                  CODEX$(JJ&) = CODEXT$(CICOMPO%)
                  CACONSU(JJ&) = CACOLO
                  IDLOT$(JJ&) = TRIM$(!IdenLote)
                    PNLD(JJ&) = "": PLCH(JJ&) = ""
                    On Error Resume Next
                  PNLD(JJ&) = !PNL_DES
                  PLCH(JJ&) = !PLC_HAS
                    On Error GoTo 0
               End If
            End If
19        .MoveNext
          Loop
       End If
       On Error GoTo 0
   End With
   '
   ' AQUI ORDENAR ASCENDENTE POR CODEX$
   For U1& = 1 To JJ&
     For U2& = U1& + 1 To JJ&
       If CODEX$(U1&) > CODEX$(U2&) Then
         SW = SWAP(CODEX$(U1&), CODEX$(U2&))
         SW = SWAP(CACONSU(U1&), CACONSU(U2&))
         SW = SWAP(IDLOT$(U1&), IDLOT$(U2&))
         SW = SWAP(PNLD$(U1&), PNLD$(U2&))
         SW = SWAP(PLCH$(U1&), PLCH$(U2&))
         KLJ = KKK
       End If
     Next U2&
   Next U1
   '
   CAITAB1% = 0
   For U& = 1 To JJ&
      CODE$ = CODEX$(U&)
      CIX0% = CODINT%(CODE$)
      '
      IDLOTE$ = IDLOT$(U&)
      CACOLO = CACONSU(U&)
      If CACOLO >= 0.05 Then
         PDES$ = PNLD(U&)
         PHAS$ = PLCH(U&)
         '
         TIPOX$ = ""
         MARCAX$ = ""
         PARTNUX$ = ""
         PROVX$ = ""
         FINGRE = 0
         FEFAB = 0
         TRACLIX$ = ""
         VERIX$ = ""
         If IDLOTE$ <> "" Then
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
              CANTID# = CACONSU(U&)
              On Error GoTo 0
              '
            End With
         End If
         '
30       CAITAB1% = CAITAB1% + 1
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
         TETABU1$(15, CAITAB1%) = CANTID#
      End If
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
        'If XVALO(Text11(KKU&)) >= 0.05 Then
          Call REPLA(TTYY$, TRIM$(Label23(KKU&)), 1, 16)
          Call REPLA(TTYY$, Text11(KKU&), 17, 16)
          Call REPLA(TTYY$, Text16(KKU&), 33, 16)
          Call REPLA(TTYY$, Text17(KKU&), 49, 16)
          TTXX$ = TTXX$ + TTYY$
        'End If
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
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Call ACTECOR(2)     ' para seleccionar solo las activas lanzadas
    Call CommSel("ECORDEP/Indice de Ordenes de Fabricación Pendientes", Command5, Label32)
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
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Command7.Enabled = False
    '
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE CantRes - CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC, Idenlote ASC"
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveLast
      If Err Then
         Call MENSERR(24, " No Existen Reservas efectuadas. ")
         GoTo 99
      End If
      FIN& = .RecordCount
      .MoveFirst
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        K& = .AbsolutePosition
        Call TRACE(20, K&, FIN&)
        NUMORFA$ = !IdSubOr
        If STATORFA%(NUMORFA$) < 3 Then
           '
           Y$ = REGBLAN("RESERVA")
           Call REPLA(Y$, MKI$(!Cod_Inte), 1, 2)
           Call REPLA(Y$, !IdSubOr, 3, 12)
           Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
           Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
           Call REPLA(Y$, MKD$(!CantRes), 19, 8)
           Call REPLA(Y$, MKD$(!CantCons), 27, 8)
           Call REPLA(Y$, MKD$(!CantRes - !CantCons), 35, 8)
           Call REPLA(Y$, !IdenLote, 43, 16)
           Call REPLA(Y$, Chr$(!Deposito), 60, 1)
           Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
           '
           JJ& = JJ& + 1
           Call GRAREG("RESERVA", Y$, JJ&)
           '
         End If
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("RESERVA", JJ&)
    Call FILESORT("RESERVA", "", 3, 4)
    '
    Call FINAL("*LIRESLOT")
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
   '
   Command8.Enabled = False
   Call TRAZABI04.Command7_Click
   Command8.Enabled = True
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
    CIXI% = CODINT%(Label27)
    Call ARMALILOT1(CIXI%, Label23(Index))
    Call FRESHECHO("!LILOTSE")
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
         CAMAXI = XVALO(Mid$(VDEV2$, 33, 12))
      End If
      Text11(Index) = FORMATNUM$(CAMAXI, "F10.1")
    End If
    '
    Text11(Index).SetFocus
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    BAUTOTRA% = 1
    Call AUTOTRA
    BAUTOTRA% = 0
    Command9.Enabled = True
End Sub

Private Sub controlmovimientosdestockrepetidos_Click()
    Call Command8_Click
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      MARCOBARRA.Top = 7300
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

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

Private Sub grabar_Click()
  Call Command13_Click
End Sub

Private Sub kardex_Click()
    Call Command10_Click
End Sub

Private Sub Label27_Change()
   CIXI% = CODINT%(Label27)
   If CIXI% > 0 Then
      Call ARMALILOT(CIXI%)
      Screen.MousePointer = 1
   End If
End Sub

Sub ARMALILOT1(CIXI%, LOTEAC$)
    '
    Static VERIRECLE%
    '
    If VERIRECLE% < 1 Then
       VERIRECLE% = 1
       If RECLEN%("LILOTSE") < 60 Then
          Call MENSERR(24, "Longitud de Registro Incorrecta\para Controlador 'LILOTSE'.\   \Reinstale el Controlador .RFS")
          Call FINAL("")
       End If
    End If
    '
    NORFA$ = Label32
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + "AND ((CanAlta - CanConsu) > 0 "    ' no trae los consumidos
    SQLX$ = SQLX$ + "OR ISNULL(CanConsu) = true or ISNULL(CanReser) = true)"
    SQLX$ = SQLX$ + "ORDER BY FechAlta ASC, IdenLote ASC "
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
           IDLOT$ = TRIM$(!IdenLote)
           '
           Call ACSALOT(CIXI%, IDLOT$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%)
           '
           TTXX$ = REBLA$
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, IDLOT$, 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F12,1"), 25, 12)
           '
           CANCON# = 0
           On Error Resume Next
           CANCON# = !CanConsu
           On Error GoTo 0
           CSALDO# = !CanAlta - CANCON#
           CSALFI# = CSALDO#
           SQLZ$ = " SELECT * FROM RESERVA "
           SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
           SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDLOT$ & "' "
           SQLZ$ = SQLZ$ + "AND IdSubOr <> '" & NORFA$ & "' "
           SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
           SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
           Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
           '
           With ReserTemp
              On Error Resume Next
              .MoveFirst
              If Err < 1 Then
                 On Error GoTo 0
                 Do While Not .EOF
                   CSALI# = !CantRes - !CantCons
                   If CSALI# >= 0.05 Then
                     CSALDO# = CSALDO# - CSALI#
                   End If
                 .MoveNext
                  Loop
              End If
           End With
           '
           For KKI% = 0 To 5
              If Label23(KKI%) = IDLOT$ Then
                 If Label23(KKI%) <> LOTEAC$ Then
                   CSALDO# = CSALDO# - XVALO(Text11(KKI%))
                 End If
              End If
           Next KKI%
           '
           If CSALDO# >= 0.05 Or CSALFI# > 0 Then
              Call REPLA(TTXX$, FORMATNUM$(CSALFI#, "F12,1"), 37, 12)
              Call REPLA(TTXX$, FORMATNUM$(CSALDO#, "F12,1"), 49, 12)
              JJ& = JJ& + 1
              Call GRAREG("!LILOTSE", TTXX$, JJ&)
           End If
         .MoveNext
         Loop
       End If
    End With
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub
'

Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
    Label17 = ""
    Label20 = ""
    Label21 = ""
    Label33 = ""
    HOII% = HOY(HOS$)
    Label24 = HOS$
    Text8.Enabled = True
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    List1.Clear
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
'    If UCase$(TRIM$(Label7)) = "AUTO" Then
'       Call CARLICONSU(1)
'       GoTo 99
'    End If
'    '
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
           Label34 = FORMATNUM$(!CanProy, "F8.1")
           Label37 = FORMATNUM$(!CanApro, "F8.1")
           Label38 = FORMATNUM$(CAPENDE#, "F8.1")
           '
         End With
         Call CARLICONSU(1)
         '
    End If
99  Screen.MousePointer = 1
End Sub

Sub CARLICONSU(MODO%)
    ' MODO%=1 DESDE LAS RESERVAS
    List1.Clear
    Dim LOTRESE$(2048), CANRESE(2048), CANTRAZA(2048)
    Label8 = (-1)
    CACONJU# = XVALO(Label38)
    If CACONJU# < 0.1 Then Exit Sub
    '
    AUTOTRAZA% = 0
    If UCase(TRIM$(Label7)) = "AUTO" Then
       If CONTROL$("TRAZABI", "MODOTRAZ") = "AUTOMA" Then
          AUTOTRAZA% = 1
       End If
    End If
    '
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
    SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC "
    '
    On Error Resume Next    ' POR SI NO ESTÁ SETEADO NuOrItem
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    If Err Then
       On Error GoTo 0
       SQLX$ = " SELECT * FROM RESERVA "
       SQLX$ = SQLX$ + "WHERE IdSubOr = '" & Label32 & "' "
       Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    End If
    On Error GoTo 0
    '
    CAIT% = 0
    With RESERTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          Do While (.EOF = False)
            CIXI% = !Cod_Inte
            TTZZ$ = Space$(64)
            IDLOCO$ = TRIM$(!IdenLote)
            CACOLO = !CantRes - !CantCons: If CACOLO < 0 Then CACOLO = 0
            CANTRA = 0
            If IDLOCO$ <> "" Then CANTRA = CACOLO
            Call REPLA(TTZZ$, IDLOCO$, 1, 16)
            Call REPLA(TTZZ$, FORMATNUM$(CACOLO, "F12.1"), 17, 16)
            On Error Resume Next
            Call REPLA(TTZZ$, !PNL_DES, 33, 16)
            Call REPLA(TTZZ$, !PLC_HAS, 49, 16)
            On Error GoTo 0
            '
            For JJ& = 1 To CAIT%
               If CIJ%(JJ&) = CIXI% Then
                  CANRESE(JJ&) = CANRESE(JJ&) + CACOLO
                  CANTRAZA(JJ&) = CANTRAZA(JJ&) + CANTRA
                  If IDLOCO$ <> "" Then
                     LOTRESE$(JJ&) = LOTRESE$(JJ&) + TTZZ$
                  End If
                  GoTo 19
               End If
            Next JJ&
            CAIT% = CAIT% + 1
            CIJ%(CAIT%) = !Cod_Inte
            USOI(CAIT%) = !UsoUnit
            CANRESE(CAIT%) = CACOLO
            CANTRAZA(CAIT%) = CANTRA
            LOTRESE$(CAIT%) = TTZZ$
19        .MoveNext
          Loop
       End If
       On Error GoTo 0
    End With
    '
    If CAIT% < 1 Then    ' No Existen Reservas efectuadas.
       ' arma composicion segun formula standard
       CODPRO$ = Label28
       Call LEEEXPLOMUL(CODPRO$, 1, 1)
       For JJ& = 1 To CAIT%
          LOTRESE$(JJ&) = ""
          CANRESE(JJ&) = CACONJU# * USOI(JJ&)
          CANTRAZA(JJ&) = 0
       Next JJ&
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
          ZZ$ = Space$(128)
          Call REPLA(ZZ$, CODEXT$(CIJ%(I1%)), 1, 16)
          Call REPLA(ZZ$, DESCRIT$(CIJ%(I1%)), 17, 40)
            NEBRU = CANRESE(I1%)
          Call REPLA(ZZ$, FORMATNUM$(NEBRU, "F13.1"), 57, 13)
            SUMATRA = CANTRAZA(I1%)
          Call REPLA(ZZ$, FORMATNUM$(SUMATRA, "F12.1"), 71, 12)
          Call REPLA(ZZ$, FORMATNUM$(USOI(I1%), "F16.6"), 97, 16)
          '
          ZZ$ = ZZ$ + LOTRESE$(I1%)
          List1.AddItem ZZ$
          '
          ' TRAZADO AUTOMATICO POR F.I.F.O. PARA DETERPLUS
          If AUTOTRAZA% > 0 Then
             If SUMATRA < NEBRU Then
                LINDEX = List1.ListCount
                List1.ListIndex = LINDEX - 1
                Call List1_Click
                Call AUTOTRA
             End If   ' FIN DE AUTOTRAZA
          End If
          '
       End If
    Next I1%
    Screen.MousePointer = 1
    '
End Sub

Sub AUTOTRA()
    '
    NEBRU = XVALO(Text4)
    SUMATRA = XVALO(Label16)
    CIIX% = CODINT%(Label27)
    '
    If CIIX% < 1 Then Exit Sub
    If List1.ListIndex < 0 Then Exit Sub
    If SUMATRA >= NEBRU Then Exit Sub
    '
    Call ARMALILOT(CIIX%)
    For KKUU& = 1 To ULTREG&("!LILOTSE")
       XX3$ = REGLEIDO$("!LILOTSE", KKUU&)
       VDEVU$ = TRIM$(Left$(XX3$, 16))
       VDEV2$ = Mid$(XX3$, 17)
       '
       If VDEVU$ <> "" Then
          CAMAXI = XVALO(Text4)
          For KKU& = 0 To 5
             CAMAXI = CAMAXI - XVALO(Text11(KKU&))
          Next KKU&
          If CAMAXI < 0 Then CAMAXI = 0
          '
          If XVALO(Mid$(VDEV2$, 33, 12)) < CAMAXI Then
             CAMAXI = XVALO(Mid$(VDEV2$, 33, 12))
          End If
          '
          If CAMAXI >= 0.05 Then
            For KKU& = 0 To 5
              If TRIM$(Label23(KKU&)) = VDEVU$ Then
                 CANLO = XVALO(Text11(KKU&)) + CAMAXI
                 CAMAXI = 0
                 Text11(KKU&) = FORMATNUM$(CANLO, "F10.1")
                 SUMATRA = SUMATRA + CAMAXI
                 GoTo 29
              End If
            Next KKU&
            '
            ' BUSCA EL PRIMERO VACÍO
            For KKU& = 0 To 5
              If TRIM$(Label23(KKU&)) = "" Then
                 Label23(KKU&) = VDEVU$
                 Text11(KKU&) = FORMATNUM$(CAMAXI, "F10.1")
                 SUMATRA = SUMATRA + CAMAXI
                 GoTo 29
              End If
            Next KKU&
          End If
       End If
29     If SUMATRA >= NEBRU Then GoTo 39
    Next KKUU&
    '
39  Call Command2_Click    ' PARA SUBIRLO
    '
End Sub
'
Private Sub Label32_DblClick()
    ' USAR PARA LIMPIAR TODO
    Label32 = ""
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
       If KKU& < 6 Then
         Label23(KKU&) = RTrim(Left$(TTYY$, 16))
         Text11(KKU&) = TRIM$(Mid$(TTYY$, 17, 16))
         Text16(KKU&) = TRIM$(Mid$(TTYY$, 33, 16))
         Text17(KKU&) = TRIM$(Mid$(TTYY$, 49, 16))
       End If
       KKU& = KKU& + 1
    Next KKL&
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

Private Sub materialreservado_Click()
    Call Command7_Click
End Sub

Private Sub printregistroytrazabilidad_Click()
   Call Command16_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
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
    
    CIXI% = CODINT%(Label27)
    If CIXI% > 0 Then
       LTXX$ = TRIM$(Label23(Index))
       If LTXX$ <> "" Then
          Call ACSALOTE(CIXI%, LTXX$)
          If BAUTOTRA% < 1 And UCase$(TRIM$(Label7)) <> "AUTO" Then
             Call ARMALILOT1(CIXI%, Label23(Index))
          End If
          Screen.MousePointer = 1
       End If
    End If
End Sub

Private Sub Text11_DblClick(Index As Integer)
    Text11(Index) = ""
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

Private Sub Text4_KeyPress(KeyAscii As Integer)

    If KeyAscii% = 13 Then
       Call Text4_LostFocus
    End If
    
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
   List1.refresh
   '
End Sub

Private Sub Text8_GotFocus()
    If UCase$(TRIM$(Label7)) = "AUTO" Then Command13.SetFocus
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub
