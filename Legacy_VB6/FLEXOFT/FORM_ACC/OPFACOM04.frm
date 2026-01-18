VERSION 5.00
Begin VB.Form OPFACOM04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Proceso de Facturación de Compras"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Redondeo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   620
      Left            =   5520
      TabIndex        =   70
      Top             =   3285
      Width           =   5250
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
         Left            =   2310
         TabIndex        =   72
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label8 
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
         Left            =   2625
         TabIndex        =   74
         Top             =   210
         Width           =   2430
      End
      Begin VB.Label Label11 
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
         TabIndex        =   73
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   0
         TabIndex        =   71
         Top             =   285
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   7475
      TabIndex        =   68
      Top             =   4725
      Width           =   1350
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Carga por Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   160
         TabIndex        =   69
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Accesos directos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5520
      TabIndex        =   56
      Top             =   105
      Width           =   5265
      Begin VB.CommandButton Command4 
         Caption         =   "C-Con"
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
         Left            =   1305
         Picture         =   "OPFACOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Base de Datos de Cuentas Contables"
         Top             =   345
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "I-B"
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
         Left            =   3225
         MaskColor       =   &H00800000&
         Picture         =   "OPFACOM04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Percepción Ingresos Brutos - Tabla de Provincias"
         Top             =   345
         Width           =   750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "C-Cos"
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
         Left            =   345
         Picture         =   "OPFACOM04.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Base de Datos de Centros de Costos"
         Top             =   345
         Width           =   750
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Tablas"
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
         Left            =   2265
         Picture         =   "OPFACOM04.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Tablas de Grupos, Familias e Items y Conceptos"
         Top             =   345
         Width           =   750
      End
      Begin VB.CommandButton Command9 
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
         Left            =   4185
         Picture         =   "OPFACOM04.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cotización Dolar y Tipo de Cambio"
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Diferencia de cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   55
      Top             =   4680
      Width           =   5295
      Begin VB.CommandButton Command3 
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
         Index           =   7
         Left            =   4800
         TabIndex        =   66
         Top             =   240
         Width           =   180
      End
      Begin VB.Label DECUENTA 
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
         Index           =   7
         Left            =   1155
         TabIndex        =   60
         Top             =   600
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   7
         Left            =   3150
         TabIndex        =   59
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor -"
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
         Left            =   600
         TabIndex        =   58
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1965
         TabIndex        =   57
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   8880
      TabIndex        =   51
      Top             =   4725
      Width           =   1890
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bi-Monetario"
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
         Left            =   210
         TabIndex        =   54
         Top             =   525
         Width           =   1600
      End
      Begin VB.OptionButton Option16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pesos"
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
         Left            =   210
         TabIndex        =   53
         Top             =   315
         Value           =   -1  'True
         Width           =   1600
      End
      Begin VB.OptionButton Option19 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Multi-Moneda"
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
         Left            =   210
         TabIndex        =   52
         Top             =   735
         Width           =   1600
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Libro IVA Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5505
      TabIndex        =   49
      Top             =   4725
      Width           =   1905
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Suprime Asien- tos de Diario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   160
         TabIndex        =   50
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Validaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5520
      TabIndex        =   47
      Top             =   3990
      Width           =   5250
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Distingue Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   67
         Top             =   210
         Width           =   1740
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Requiere Ingreso  C.A.I."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   420
         TabIndex        =   48
         Top             =   210
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   6720
      Left            =   10920
      ScaleHeight     =   6660
      ScaleWidth      =   915
      TabIndex        =   41
      Top             =   0
      Width           =   975
      Begin VB.PictureBox Picture16 
         Height          =   5895
         Left            =   0
         ScaleHeight     =   5835
         ScaleWidth      =   1005
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   5850
            Left            =   0
            Picture         =   "OPFACOM04.frx":09F2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   45
         Top             =   4935
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
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
         Left            =   140
         Picture         =   "OPFACOM04.frx":2D8A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ayuda en Linea FLEXOFT"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   140
         Picture         =   "OPFACOM04.frx":3094
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Guardar Configuración"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Left            =   140
         Picture         =   "OPFACOM04.frx":339E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "OPFACOM04.frx":34E8
         Top             =   5300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tasas de IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5265
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   36
         Top             =   3675
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Index           =   3
         Left            =   1155
         TabIndex        =   22
         Text            =   " "
         Top             =   2835
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   19
         Top             =   2835
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Index           =   2
         Left            =   1155
         TabIndex        =   16
         Text            =   " "
         Top             =   1995
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   13
         Top             =   1995
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Index           =   1
         Left            =   1155
         TabIndex        =   10
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   7
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox TASIVA 
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
         Index           =   0
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4830
         TabIndex        =   1
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1995
         TabIndex        =   40
         Top             =   3750
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A. Aduana -"
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
         Left            =   630
         TabIndex        =   39
         Top             =   3750
         Width           =   1695
      End
      Begin VB.Label CUEIVA 
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
         Index           =   4
         Left            =   3150
         TabIndex        =   38
         Top             =   3675
         Width           =   1695
      End
      Begin VB.Label DECUENTA 
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
         Index           =   4
         Left            =   1155
         TabIndex        =   37
         Top             =   4035
         Width           =   3870
      End
      Begin VB.Label DECUENTA 
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
         Index           =   3
         Left            =   1155
         TabIndex        =   24
         Top             =   3195
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   3
         Left            =   3150
         TabIndex        =   23
         Top             =   2835
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-3:"
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
         TabIndex        =   21
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1995
         TabIndex        =   20
         Top             =   2910
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Index           =   2
         Left            =   1155
         TabIndex        =   18
         Top             =   2355
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   2
         Left            =   3150
         TabIndex        =   17
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-2:"
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
         TabIndex        =   15
         Top             =   2070
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1995
         TabIndex        =   14
         Top             =   2070
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Index           =   1
         Left            =   1155
         TabIndex        =   12
         Top             =   1515
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   1
         Left            =   3150
         TabIndex        =   11
         Top             =   1155
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TASA-1:"
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
         Left            =   105
         TabIndex        =   9
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1995
         TabIndex        =   8
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label DECUENTA 
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
         Index           =   0
         Left            =   1155
         TabIndex        =   6
         Top             =   675
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   0
         Left            =   3150
         TabIndex        =   5
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Basica:"
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
         TabIndex        =   3
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1995
         TabIndex        =   2
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cuentas Contables Percepciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1940
      Index           =   1
      Left            =   5520
      TabIndex        =   25
      Top             =   1335
      Width           =   5265
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
         Height          =   285
         Index           =   6
         Left            =   4830
         TabIndex        =   27
         Top             =   1125
         Width           =   175
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
         Height          =   285
         Index           =   5
         Left            =   4830
         TabIndex        =   26
         Top             =   330
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Index           =   7
         Left            =   1995
         TabIndex        =   35
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percepcion IVA - "
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
         Index           =   7
         Left            =   -210
         TabIndex        =   34
         Top             =   405
         Width           =   2595
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Index           =   6
         Left            =   1995
         TabIndex        =   33
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percepción Ganancias - "
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
         Index           =   6
         Left            =   -210
         TabIndex        =   32
         Top             =   1260
         Width           =   2595
      End
      Begin VB.Label CUEIVA 
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
         Index           =   6
         Left            =   3150
         TabIndex        =   31
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label DECUENTA 
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
         Index           =   6
         Left            =   1155
         TabIndex        =   30
         Top             =   1485
         Width           =   3870
      End
      Begin VB.Label CUEIVA 
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
         Index           =   5
         Left            =   3150
         TabIndex        =   29
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label DECUENTA 
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
         Left            =   1155
         TabIndex        =   28
         Top             =   690
         Width           =   3870
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPFACOM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command14_Click()
   Call CommSel("ECUECON", Command14, Label11)
End Sub

Private Sub Command2_Click()
    '
    For IXX% = 0 To 7
      TXXX$ = TRIM$(Str$(IXX%))
      Call GRACONTROL$("", "CUEIVA-" + TXXX$, TRIM$(CUEIVA(IXX%).Caption))
      If IXX% < 4 Then
         TASAX$ = REPLACAR$(TASIVA(IXX%).TEXT, ",", ".")
         Call GRACONTROL$("", "TASIVA-" + TXXX$, TASAX$)
         TASA = XVALO(TASAX$)
         On Error Resume Next
         If TASA > 0.005 Then
             FORFACOM.TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
             FORFAC00.TASIVA(IXX%).Caption = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
           Else
             FORFACOM.TASIVA(IXX%).Caption = ""
             FORFAC00.TASIVA(IXX%).Caption = ""
         End If
         On Error GoTo 0
      End If
    Next IXX%
    If Check1.Value Then
      Call GRACONTROL("", "VALICAI", "SI")
    Else
      Call GRACONTROL("", "VALICAI", "NO")
    End If
    'OPCION DE DISTINGUIR POR LETRA O NO
    '
    COMPALET$ = "NO"
    If Check3.Value = 1 Then COMPALET$ = "SI"
    Call GRACONTROL("FACOMPRA", "VERILEFA", COMPALET$)
    
    '\\\OT-06-0208-WAR-05/06/06///
    If Check2.Value Then
      Call GRACONTROL("LIVACOM", "SUPRIVA", "SI")
    Else
      Call GRACONTROL("LIVACOM", "SUPRIVA", "NO")
    End If
    '\\\OT-06-0208-WAR-FIN///
    '
    '\\\OT-08-0162-OM-15/04/08///
    'GRABA SELECCION DE TILDAR POR DETALLE EN FACTURACION DE PROVEEDORES
    '
    VALOR$ = "NO"
    If Check4.Value = 1 Then VALOR$ = "SI"
    Call GRACONTROL("FACOMPRA", "CARGADET", VALOR$)
    '\\\OT-08-0162-OM-FIN///
    '
    MODOMONE$ = "PESOS"
    If Option15.Value = True Then MODOMONE$ = "BIMONETA"
    If Option19.Value = True Then MODOMONE$ = "MULTIMON"
    Call GRACONTROL("GESACRED", "MODOMONE", MODOMONE$)
    '
    IMPURED$ = TRIM$(Label11)
    Call GRACONTROL("FACOMPRA", "REDONDEO", IMPURED$)
    Hide
    
    '
End Sub

Private Sub Command3_Click(Index As Integer)
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        CUEIVA(Index).Caption = VDEVU$
        DECUENTA(Index).Caption = VALACT2$("ECUECON")
    End If
End Sub
'

Sub Command4_Click()
    Command4.Enabled = False
    If EXISTE%("CONTAB04.EXE") > 0 Then
        Call COMUNI("Actualice la Base de Datos de Cuentas Contables\a Través del Plan de Cuentas\del Sistema de Contabilidad General.\  \Pulse 'Aceptar' para Pasar a este Sistema.")
        Call CONECRUN("CONTAB04", "Contabilidad General - Caja y Bancos")
      Else
        Call CARCUEINGAS
    End If
    Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   Call CARTAPROVIN
   Command5.Enabled = True
End Sub

Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    TO1& = ULTREG("TACENCOS") '\\\OT-06-0244-WAR-28/06/06///
    Call TRALOCAL("TACENCOS", "")
10  VTB% = VENTABU("TACENCOS")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACENCOS")
        X$ = REGLEIDO$("!TACENCOS", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 8)))
        Call REPLA(X$, CD$, 1, 8)
        Call GRAREG("!TACENCOS", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(CD$) < 2 Then
          Call MENSERR(24, "Código de Centro Costos Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que el CC posea descripcion
        DESC$ = TRIM$(Mid$(X$, 9, 40))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Centro de Costos sin Descripción")
          GoTo 10
        End If
        ' Valida que el CC no este repetido
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TACENCOS", U2&)
          If TRIM$(Mid$(X$, 1, 8)) = TRIM$(Mid$(Y$, 1, 8)) Then
            Call MENSERR(24, "Código de Centro de Costos Repetido")
            GoTo 10
          End If
        Next U2&
        '
      Next U&
      Call TRACENTRAL("TACENCOS", "")
      '\\\OT-06-0244-WAR-28/06/06///
      TO2& = ULTREG&("TACENCOS")
      If TO2& = TO1& + 1 Then 'Si Agrego un registo
        'Le coloco la fecha 'desde' de la Vigencia
        HOII% = HOY(HO$)
        XPLA$ = REGLEIDO$("TACENCOS", TO2&)
        Call REPLA(XPLA$, MKI$(HOII%), 45, 2)
        Call GRAREG("TACENCOS", XPLA$, TO2&)
      End If
      '\\\OT-06-0244-WAR-FIN///
    End If
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("ACRE_GES")
End Sub

Sub Command8_Click()
    '
      Command8.Enabled = False
    '
    OPX% = COMALTER%("Seleccione la Tabla:\\Grupos\Familias\Conceptos")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    '
    Select Case OPX%
        Case 1 'GRUPOS
10          Call TRALOCAL("TABGRUP", "")
            VTB% = VENTABU("TABGRUP")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TABGRUP")
                X$ = REGLEIDO$("!TABGRUP", U&)
                CD$ = TRIM((Mid$(X$, 1, 6)))
                Call REPLA(X$, CD$, 1, 6)
                Call GRAREG("TABGRUP", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If TRIM(CD$) = "" Then
                  Call MENSERR(24, "Código de Grupo no Válido.")
                  GoTo 10
                End If
                ' Valida que la familia posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 48))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Grupo sin Descripción")
                  GoTo 10
                End If
                'valida que haya ingresado proratera
                PRO$ = TRIM$(Mid$(X$, 55, 1))
                If Len(PRO$) < 1 Then
                  Call MENSERR(24, "No ha Definidio el Campo Prorratea.")
                  GoTo 10
                End If
                '
              Next U&
              Call FILESORT("TABGRUP", "TABGRUP", 1, 1, "ASC")
              Call TRACENTRAL("TABGRUP", "")
            End If
        Case 2 'FAMILIAS
20          Call TRALOCAL("TABFAMI", "")
            VTB% = VENTABU("TABFAMI")
            If VTB% >= 0 Then
              For U& = 1 To ULTREG&("!TABFAMI")
                X$ = REGLEIDO$("!TABFAMI", U&)
                CD$ = TRIM$((Mid$(X$, 1, 6)))
                Call REPLA(X$, CD$, 1, 6)
                Call GRAREG("TABFAMI", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If TRIM(CD$) = "" Then
                  Call MENSERR(24, "Código de Familia no Válido.")
                  GoTo 20
                End If
                ' Valida que la familia posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 48))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Familia sin Descripción")
                  GoTo 20
                End If
                'valida que haya ingresado el grupo de pertenencia
                NG$ = TRIM((Mid$(X$, 55, 6)))
                If TRIM(NG$) <= "" Then
                  Call MENSERR(24, "No ha Definidio el Grupo de Pertenencia.")
                  GoTo 20
                End If
                '
                'valida que el codigo no sea repetido
                For U2& = 1 To U& - 1
                  Y$ = REGLEIDO$("!TABFAMI", U2&)
                        If TRIM$(Mid$(X$, 1, 6)) = TRIM$(Mid$(Y$, 1, 6)) Then
                            Call MENSERR(24, "Código de Item Repetido")
                            GoTo 20
                        End If
                Next U2&
                '
              Next U&
              Call FILESORT("TABFAMI", "TABFAMI", 1, 1, "ASC")
              Call TRACENTRAL("TABFAMI", "")
            End If
        Case 3 'ITEMS Y CONCEPTOS.
30          Call TRALOCAL("TABITCO", "")
            VTB% = VENTABU("TABITCO")
            If VTB% >= 0 Then
                For U& = 1 To ULTREG&("!TABITCO")
                X$ = REGLEIDO$("!TABITCO", U&)
                CD1$ = TRIM$(Mid$(X$, 1, 6))
                Call REPLA(X$, CD1$, 1, 6)
                Call GRAREG("TABITCO", X$, U&)
                '
                ' ValidA QUE EL CODIGO SEA VALIDO
                If Len(CD1$) <= 0 Or Len(CD1$) > 6 Then
                  Call MENSERR(24, "Código de Item no Válido.")
                  GoTo 30
                End If
                'valida que el codigo no sea repetido
                For U2& = 1 To U& - 1
                  Y$ = REGLEIDO$("!TABITCO", U2&)
                        If TRIM$(Mid$(X$, 1, 6)) = TRIM$(Mid$(Y$, 1, 6)) Then
                            Call MENSERR(24, "Código de Item Repetido")
                            GoTo 30
                        End If
                Next U2&
                ' Valida que EL ITEM posea descripcion
                DESC$ = TRIM$(Mid$(X$, 7, 32))
                If Len(DESC$) < 1 Then
                  Call MENSERR(24, "Código de Item sin Descripción")
                  GoTo 30
                End If
                'valida que haya familia de pertenencia
                NF$ = TRIM(Mid$(X$, 41, 6))
                If TRIM(NF$) = "" Then
                  Call MENSERR(24, "No ha Definidio la Familia de Pertenencia.")
                  GoTo 30
                End If
                'NO VALIDA LA UNIDAD DE MEDIDA XQ
                'ALGUNOS TIENEN Y OTROS NO
              Next U&
              Call FILESORT("TABITCO", "TABITCO", 1, 1, "ASC")
              Call TRACENTRAL("TABITCO", "")
            End If
    End Select
    '
99  Command8.Enabled = True
    '
    '
End Sub

Sub Command9_Click()
   '\\\OT-06-0023-FG-26/01/06
   Command9.Enabled = False
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     GoTo 99
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
99 Command9.Enabled = True
   '\\\OT-06-0023-FG-FIN
End Sub

Private Sub CUEIVA_Change(Index As Integer)
   DECUENTA(Index).Caption = DENOCUE$(CUEINT%(CUEIVA(Index).Caption))
End Sub

Private Sub Form_Load()
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame10(0).BackColor = &H8000000F
      Me.Frame10(1).BackColor = &H8000000F
      Me.Frame15.BackColor = &H8000000F
      Me.Option15.BackColor = &H8000000F
      Me.Option16.BackColor = &H8000000F
      Me.Option19.BackColor = &H8000000F
      Me.Check1.BackColor = &H8000000F
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      MARCOBARRA.Top = 5450
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If
    
    For IXX% = 0 To 7
      TXXX$ = TRIM$(Str$(IXX%))
      CUEIVA(IXX%).Caption = Control$("", "CUEIVA-" + TXXX$)
      If CUEINT%(CUEIVA(IXX%)) < 1 Then
        CUEIVA(IXX%) = TRIM$(CUEIVA(IXX%)) + "-inex"
      End If
      If IXX% < 4 Then
         TASA = XVALO(Control$("", "TASIVA-" + TXXX$))
         TASIVA(IXX%).TEXT = CSTRING$(MKS$(TASA), 3, 6, 2, 2)
      End If
    Next IXX%
    '
    If Control$("", "VALICAI") = "SI" Then
        Check1.Value = 1
    End If
    '\\\OT-06-0208-WAR-05/06/06///
    If Control$("LIVACOM", "SUPRIVA") = "SI" Then
      Check2.Value = 1
    End If
    '\\\OT-06-0208-WAR-FIN///
    If Control$("GESACRED", "MODOMONE") = "BIMONETA" Then Option15.Value = True
    If Control$("GESACRED", "MODOMONE") = "MULTIMON" Then Option19.Value = True
    '
    Check3.Value = 0
    If Control$("FACOMPRA", "VERILEFA") = "SI" Then Check3.Value = 1
    
    '\\\OT-08-0162-OM-15/04/08///
    'LEVANTA DEL SETUP SI ESTA SELECCIONADO CARGA POR DETALLE
    '
    Check4.Value = 0
    If Control$("FACOMPRA", "CARGADET") = "SI" Then Check4.Value = 1
    '\\\OT-08-0162-OM-FIN///
    '
    
    Label11.Caption = Control$("FACOMPRA", "REDONDEO")
End Sub

Private Sub grabar_Click()
    Call Command2_Click
End Sub

Private Sub Label11_Change()
  Label8.Caption = DENOCUE$(CUEINT%(Label11.Caption))
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub
