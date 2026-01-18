VERSION 5.00
Begin VB.Form MODIPRO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación de Datos Cuenta de Proveedor"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Gan"
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
      Left            =   9870
      TabIndex        =   59
      ToolTipText     =   "Datos Retención Ganancias Proveedor Activo"
      Top             =   3510
      Width           =   540
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   330
      Left            =   9765
      TabIndex        =   57
      Top             =   4380
      Width           =   750
      Begin VB.CommandButton Command22 
         Caption         =   "Cfg"
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
         Left            =   105
         TabIndex        =   58
         ToolTipText     =   "Datos Retención I.Brutos Proveedor Activo"
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   5565
      Picture         =   "MODIPRO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Configuración - Valores Predeterminados por Omisión"
      Top             =   5985
      Width           =   855
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
      Height          =   360
      Index           =   5
      Left            =   10250
      TabIndex        =   55
      Top             =   4935
      Width           =   175
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
      Height          =   360
      Index           =   4
      Left            =   2520
      TabIndex        =   54
      Top             =   5355
      Width           =   175
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
      Height          =   360
      Index           =   3
      Left            =   2520
      TabIndex        =   53
      Top             =   4935
      Width           =   175
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
      Height          =   360
      Index           =   2
      Left            =   6720
      TabIndex        =   52
      Top             =   3890
      Width           =   175
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
      Height          =   360
      Index           =   1
      Left            =   6720
      TabIndex        =   51
      Top             =   3045
      Width           =   175
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
      Height          =   360
      Index           =   0
      Left            =   2415
      TabIndex        =   50
      Top             =   2310
      Width           =   175
   End
   Begin VB.CommandButton BOSELCUE 
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
      Left            =   10250
      TabIndex        =   49
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox DETEXT 
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
      Left            =   6510
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5355
      Width           =   3930
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   5
      Left            =   8190
      TabIndex        =   47
      Text            =   " "
      Top             =   4935
      Width           =   2250
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Imprimir"
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
      Left            =   4620
      Picture         =   "MODIPRO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5985
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Anotaciones"
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
      Left            =   2730
      Picture         =   "MODIPRO.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5985
      Width           =   1800
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   3
      Left            =   1995
      TabIndex        =   18
      Text            =   " "
      Top             =   4935
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
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
      Left            =   6405
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   3510
      Width           =   3400
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   1
      Left            =   6405
      TabIndex        =   12
      Text            =   " "
      Top             =   3045
      Width           =   360
   End
   Begin VB.TextBox DETEXT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2730
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   2310
      Width           =   3930
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   0
      Left            =   1995
      TabIndex        =   6
      Text            =   "  "
      Top             =   2310
      Width           =   465
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
      Height          =   360
      Left            =   1995
      TabIndex        =   2
      Text            =   " "
      Top             =   735
      Width           =   8445
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
      Left            =   9135
      TabIndex        =   0
      Text            =   " "
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
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
      Left            =   8610
      Picture         =   "MODIPRO.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5985
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
      Left            =   7665
      Picture         =   "MODIPRO.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5985
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
      Left            =   9555
      Picture         =   "MODIPRO.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5985
      Width           =   855
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
      Left            =   6720
      Picture         =   "MODIPRO.frx":13DC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5985
      Width           =   855
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
      Left            =   1995
      TabIndex        =   1
      Text            =   " "
      Top             =   210
      Width           =   2565
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
      Left            =   1995
      TabIndex        =   3
      Text            =   " "
      Top             =   1260
      Width           =   8445
   End
   Begin VB.TextBox Text5 
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
      Left            =   1995
      TabIndex        =   4
      Text            =   " "
      Top             =   1785
      Width           =   1095
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
      Left            =   4620
      TabIndex        =   5
      Text            =   " "
      Top             =   1785
      Width           =   5820
   End
   Begin VB.TextBox Text9 
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
      TabIndex        =   8
      Text            =   " "
      Top             =   2310
      Width           =   2565
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
      Left            =   1995
      TabIndex        =   9
      Text            =   " "
      Top             =   3045
      Width           =   2775
   End
   Begin VB.TextBox Text11 
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
      Left            =   1995
      TabIndex        =   10
      Text            =   " "
      Top             =   3570
      Width           =   2775
   End
   Begin VB.TextBox TEXT12 
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
      Left            =   1995
      TabIndex        =   11
      Text            =   " "
      Top             =   4095
      Width           =   2775
   End
   Begin VB.TextBox Text15 
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
      TabIndex        =   14
      Text            =   " "
      Top             =   3045
      Width           =   2565
   End
   Begin VB.TextBox DETEXT 
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
      Left            =   6405
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4410
      Width           =   4035
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   2
      Left            =   6405
      TabIndex        =   15
      Text            =   " "
      Top             =   3890
      Width           =   360
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
      Left            =   7875
      TabIndex        =   17
      Text            =   " "
      Top             =   3890
      Width           =   2565
   End
   Begin VB.TextBox COTEXT 
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
      Index           =   4
      Left            =   1995
      TabIndex        =   20
      Text            =   " "
      Top             =   5355
      Width           =   570
   End
   Begin VB.TextBox DETEXT 
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
      Left            =   2835
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4935
      Width           =   3405
   End
   Begin VB.TextBox DETEXT 
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
      Left            =   2835
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5355
      Width           =   3405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   230
      TabIndex        =   27
      Top             =   5985
      Width           =   2325
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   850
         Left            =   0
         Picture         =   "MODIPRO.frx":16E6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10000
      End
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cuenta:"
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
      Left            =   210
      TabIndex        =   45
      Top             =   5040
      Width           =   1710
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA"
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
      Left            =   4935
      TabIndex        =   44
      Top             =   3150
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
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
      Left            =   735
      TabIndex        =   43
      Top             =   2415
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razón Social:"
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
      Left            =   0
      TabIndex        =   42
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Cuenta:"
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
      Left            =   6615
      TabIndex        =   41
      Top             =   315
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Fantasia:"
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
      TabIndex        =   40
      Top             =   315
      Width           =   1320
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calle y Número:"
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
      Left            =   0
      TabIndex        =   39
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo.Postal:"
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
      Left            =   0
      TabIndex        =   38
      Top             =   1890
      Width           =   1905
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      TabIndex        =   37
      Top             =   1890
      Width           =   1905
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "País:"
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
      Left            =   5880
      TabIndex        =   36
      Top             =   2415
      Width           =   1905
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   0
      TabIndex        =   35
      Top             =   3150
      Width           =   1905
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
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
      Left            =   0
      TabIndex        =   34
      Top             =   3675
      Width           =   1905
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
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
      Left            =   0
      TabIndex        =   33
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT:"
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
      Left            =   5880
      TabIndex        =   32
      Top             =   3150
      Width           =   1905
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría I-B"
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
      Left            =   4935
      TabIndex        =   31
      Top             =   4040
      Width           =   1380
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro I-B:"
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
      Left            =   5880
      TabIndex        =   30
      Top             =   4040
      Width           =   1905
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cond.Pago:"
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
      Left            =   210
      TabIndex        =   29
      Top             =   5460
      Width           =   1710
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Madre:"
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
      Left            =   6195
      TabIndex        =   28
      Top             =   5040
      Width           =   1905
   End
End
Attribute VB_Name = "MODIPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GENDAREGAPRO()
'\\\OT-06-0559-RG-18/12/06///
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARGAPRO")
    Call REPLA(REBLA$, USERNAME$, 36, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 56, 2)
    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 58, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        RECORD$ = REBLA$
        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
        Call REPLA(RECORD$, Chr$(2), 3, 1)
        Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NCLIX%), RECORD$)
      End If
    Next U&
    Call CIERRARCH("DARGAPRO")
    Screen.MousePointer = 1
    '
'\\\OT-06-0559-RG-FIN///
End Sub

Private Sub Command1_Click()
    Call GRADATPRO
End Sub

Private Sub Command2_Click()
    '
    If MODO% = 1 Then
        Call CLEARFORMPRO
        NC% = 0
        Text1.TEXT = ""
        MODIPRO.Refresh
        MODIPRO.Command22.Refresh
        Text1.SetFocus
        Exit Sub
    End If
    '
    If NC% < 0 Then
        Call CARDATPRO(NC%)
        MODIPRO.Refresh
        MODIPRO.Command22.Refresh
      Else
        Text1.SetFocus
    End If
End Sub

Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   If EXISTE%(LUDAT$ + "DARIBPRO.RFS") < 1 Then
     Call COPYSTRU("DAPIBCLI", "DARIBPRO")
   End If
   '
   If ULTREG&("DARIBPRO") < 1 Then
     Call GENDARIBPRO
   End If
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("PROVED1", MKI$(NCLIE%)) <> "" Then
       If TICUEPRO%((-1) * NCLIE%) = 0 Then ' solo cuentas corrientes
         '
         If ULTREG&("TACATIBP") < 5 Then
           Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
           Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
           Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
           Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
           Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
           Call CIERRARCH("TACATIBP")
         End If
         '
         RECORD$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NCLIE%))
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         '
10       OBX$ = OBJPLA$("CARDARETIB", RECORD$)
         If OBX$ <> "" Then
           RECORD$ = OBX$
           CATIB1% = Asc(Mid$(RECORD$, 3, 1))
           If ECOARCH$("TACATIBP", Chr$(CATIB1%)) = "" Then
             Call MENSERR(24, "Mal Categoría I.Brutos")
             GoTo 10
           End If
           If CATIB1% < 1 Then
             If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
               Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
               GoTo 10
             End If
           End If
           If CATIB1% > 0 Then
             Call REPLA(RECORD$, Space$(32), 9, 32)
           End If
           PORCONVE = CVS(Mid$(RECORD$, 5, 4))
           If PORCONVE < 0 Or PORCONVE > 100 Then
             Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
             GoTo 10
           End If
           '
           Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
           Call REPLA(RECORD$, USERNAME$, 41, 20)
           Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
           Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
           Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NCLIE%), RECORD$)
           Call CIERRARCH("DARIBPRO")
           '
         End If
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    '
    If Abs(VERICUIT%) <> 1 Then
       VERICUIT% = 1
       If CONTROL$("AMAPRO", "VERICUIT") = "NO" Then
          VERICUIT% = (-1)
       End If
    End If
    '
    Call SORTCLI(2)
    '
    Screen.MousePointer = 11
    Call GRACREDOR
    If ULTREG&("PROVED1") > 0 Then
        XXX$ = REGLEIDO$("PROVED1", 1)
        Call GRAREG("PROVED1", XXX$, 1)
        Call CIERRARCH("PROVED1")
    End If
    Screen.MousePointer = 1
    '
    Call FRESHECHO("PROVED1")
    On Error Resume Next
    FORMUSEL.Hide
    On Error GoTo 0
    MODIPRO.Text1.SetFocus
    MODIPRO.Hide
    AMAPRO1.Show
    Call FRESHALL
End Sub

Private Sub Command4_Click()
    Call HELPONLINE("AMAPRO")
End Sub
'

Private Sub Command5_Click()
    NCLI$ = TRIM$(Text1.TEXT)
    If NCLI$ <> "" Then
        Command5.Enabled = False
        Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Proveedor")
    End If
End Sub

Private Sub Command6_Click()
    If NC% < 0 Then
       PRINTFORM
    End If
End Sub
'

Private Sub Command7_Click()
    Command7.Enabled = False
    OPALTPRO.Show 1
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
'\\\OT-06-0559-RG-18/12/06///
      Command8.Enabled = False
   '
   If ULTREG&("DARGAPRO") < 1 Then
     Call GENDAREGAPRO
   End If
   '
   NPRO% = XVALO(TRIM$(Text1.TEXT))
   If NPRO% > 0 Then
     If ECOARCH$("PROVED1", MKI$(NPRO%)) <> "" Then
       '
       If ULTREG&("TACRETG") < 1 Then
         Call GRAREG("TACRETG", Chr$(0) + "Exento ", 1)
         Call GRAREG("TACRETG", Chr$(1) + "Profesionales", 2)
         Call GRAREG("TACRETG", Chr$(1) + "Bienos de Uso y de Cambio", 3)
         Call SETULTREG("TACRETG", 3)
       End If
       '
       CATMAX% = ULTREG&("TACRETG")
       RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NPRO%))
       Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Then
         Call REPLA(RECORD$, Chr$(2), 3, 1)
         Call REPLA(RECORD$, Space$(32), 4, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAREGAN", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > CATMAX% Then
           Call MENSERR(24, "Mal Categoría de Ret. de Ganancias")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 4, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         '
         If CATIB1% > 0 Then
           Call REPLA(RECORD$, Space$(32), 4, 32)
         End If
         '
         Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 36, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 56, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 58, 2)
         Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NPRO%), RECORD$)
         Call CIERRARCH("DARGAPRO")
         '
       End If
     End If
   End If
   Command8.Enabled = True
'\\\OT-06-0559-RG-FIN///
End Sub


Private Sub Form_Activate()
    '
    Static MODANT%, YACTI%
    '
    If MODO% <> MODANT% Then
        YACTI% = 0
        MODANT% = MODO%
        If MODO% = 1 Then
             MODIPRO.BackColor = &HC0FFFF
             MODIPRO.Caption = "Alta de Nueva Cuenta de Proveedor"
             MODIPRO.BOSELCUE.Visible = False
          Else
             MODIPRO.BackColor = &HC0FFC0
             MODIPRO.Caption = "Modificación de Datos Cuenta de Proveedor"
             MODIPRO.BOSELCUE.Visible = True
             MODIPRO.Text1.TEXT = ""
        End If
    End If
    DoEvents
    '
    If YACTI% = 0 Then
       Text1.SetFocus
       YACTI% = 1
    End If
End Sub
'

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    If InStr(UCase$(EPAC$), "SABO") > 0 Or InStr(UCase$(EPAC$), "DOSIVA") > 0 Then
      Frame2.Visible = True
    End If
End Sub

Private Sub Text1_Click()
    NC% = 0
    Text1.SetFocus
End Sub

Private Sub Text1_DblClick()
    If MODO% = 2 Then
        Text1.TEXT = ""
        Call SELECHO("PROVED1")
        Exit Sub
      Else
        Text1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub BOSELCUE_CLICK()
    If MODO% = 2 Then
        Text1.TEXT = ""
        Call SELECHO("PROVED1")
        Text1.SetFocus
        Exit Sub
      Else
        Text1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus()
    '
    If MODO% = 1 Then
       If TRIM$(Text1.TEXT) = "" Then
          Screen.MousePointer = 11
          For Index = 0 To 5
             VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
             COTEXT(Index) = "": DETEXT(Index) = ""
             If VAPRE$ <> "" Then
                TV11% = TIPFIELD(ARCHE$(Index), 1)
                LF11% = LENFIELD(ARCHE$(Index), 1)
                ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
                ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
                If ECO$ <> "" Then
                    COTEXT(Index).TEXT = VAPRE$
                    DETEXT(Index).TEXT = ECO$
                  Else
                    COTEXT(Index).TEXT = ""
                    DETEXT(Index).TEXT = ""
                End If
             End If
             DETEXT(Index).TabStop = False
          Next Index
          '
          NPROMAX% = 0: ATX% = 0
          For U& = 1 To ULTREG&("PROVED1")
             D1$ = REGLEIDO$("PROVED1", U&)
             D2$ = REGLEIDO$("PROVED2", U&)
             NLEY% = CVI(Left$(D1$, 2))
             If NLEY% = 8224 Then
                D1$ = Mid$(D1$, 3)
                Call GRAREG("PROVED1", D1$, U&)
             End If
             'If Asc(Mid$(D2$, 119, 1)) = 0 Then
             NPX1% = CVI(D1$)
             If NPX1% > NPROMAX% Then
               If TICUEPRO%((-1) * NPX1%) = 0 Then
                   If NPX1% < 9999 Then
                      NPROMAX% = NPX1%
                   End If
                End If
             End If
          Next U&
          '
10        If NPROMAX% = 8223 Then
             NPROMAX% = 8224
          End If
          If REGICLI&((-1) * (NPROMAX% + 1)) > 0 Then
            NPROMAX% = NPROMAX% + 1
            GoTo 10
          End If
          '
          If NPROMAX% < 9999 Then
               MODIPRO.Text1.TEXT = TRIM$(Str$(1 + NPROMAX%))
             Else
               MODIPRO.Text1.TEXT = ""
          End If
          If ATX% > 0 Then
             Call SORTCLI(2)
             Call FINAL("")
          End If
          Screen.MousePointer = 1
       End If
       Exit Sub
    End If
    '
    If Text1.TEXT = "" Then
        Text1.TEXT = VALACT1$("PROVED1")
        Text1.Refresh
        NC% = (-1) * Val(Text1.TEXT)
        If NC% < 0 Then
            Call CARDATPRO(NC%)
            MODIPRO.Refresh
            MODIPRO.Command22.Refresh
            Text2.SetFocus
            Exit Sub
        End If
    End If
    '
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VALICLI% = 0
        If Val(Text1.TEXT) < 1 Or Val(Text1.TEXT) > 32767 Then
            NC% = 0
            Text1.TEXT = ""
            Text1.SetFocus
            Exit Sub
        End If
        '
        NC% = (-1) * Val(Text1.TEXT)
        '
        If MODO% = 1 Then
            If REGICLI&(NC%) > 0 Then
                Call MENSERR(24, "Cuenta Existente en la Base de Datos")
                NC% = 0
                Text1.TEXT = ""
                Text1.SetFocus
                Exit Sub
              Else
                Text2.SetFocus
                Exit Sub
            End If
        End If
        '
        If REGICLI&(NC%) < 1 Then
            Call MENSERR(24, "La Cuenta" + Str$(Abs(NC%)) + " Invocada no Figura\en la Base de Datos de PROVEEDORES")
            NC% = 0
            Text1.SetFocus
            Exit Sub
        End If
        '
        If NC% < 0 Then
              Call CARDATPRO(NC%)
              MODIPRO.Refresh
              MODIPRO.Command22.Refresh
              Text2.SetFocus
              Exit Sub
           Else
              Text1.SetFocus
        End If
    End If
End Sub
'

Private Sub Text8_Click()
    Text7.SetFocus
End Sub
'

Private Sub Text14_CLICK()
    Text13.SetFocus
End Sub
'

Private Sub Text17_CLICK()
    Text16.SetFocus
End Sub
'

Private Sub Text20_CLICK()
    Text19.SetFocus
End Sub
'
Private Sub Text22_CLICK()
    Text21.SetFocus
End Sub
'
Private Sub Text24_CLICK()
    Text23.SetFocus
End Sub
'
Private Sub Text26_CLICK()
    Text25.SetFocus
End Sub
'

Private Sub Text28_CLICK()
    Text27.SetFocus
End Sub
'

Private Sub text1_lostfocus()
    '
    If MODO% = 1 Then
        '
        NC% = (-1) * Val(Text1.TEXT)
        If REGICLI&(NC%) > 0 Then
             Call MENSERR(24, "Cuenta Existente en la Base de Datos")
             NC% = 0
             Text1.TEXT = ""
             On Error Resume Next
             Text1.SetFocus
             On Error GoTo 0
             Exit Sub
           Else
             Exit Sub
        End If
    End If
    '
End Sub

Private Sub Text15_LostFocus()
   NC1% = (-1) * Val(Text1.TEXT)
   CUTT$ = TRIM$(Text15.TEXT)
   If CUTT$ = "" Then Exit Sub
   '
   If VALICUIT%(CUTT$) <> 1 Then
      Call MENSERR(24, "Número de C.U.I.T. no Válido")
      Text15.SetFocus
   End If
   NPP% = REPICUIT%(CUTT$, NC1%)
   If NPP% <> 0 Then
      TTXX$ = "Número de C.U.I.T. Repetido.\  \Corresponde a " + TRIM$(RASOCLI$(NPP%))
      If COMALTER%(TTXX$ + "\\Cancelar Alta\Registrar Igual") = 2 Then
        Exit Sub
      End If
      Text15.TEXT = ""
      Text15.SetFocus
   End If
End Sub

Private Sub Text2_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text3_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text4_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text5_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text6_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text7_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text9_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text10_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text11_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text12_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text13_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub Text15_GotFocus()
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub COTEXT_DblClick(Index As Integer)
    If NC% > (-1) Then
        Text1.SetFocus
        Exit Sub
      Else
        Call SELECHO(ARCHE$(Index))
        If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
            COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
            COTEXT(Index).Refresh
            DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
            DETEXT(Index).Refresh
            DETEXT(Index).TabStop = False
        End If
        COTEXT(Index).SetFocus
    End If
End Sub
'
Private Sub BOTSEL_Click(Index As Integer)
    If NC% > (-1) Then
        Text1.SetFocus
        Exit Sub
      Else
        Call SELECHO(ARCHE$(Index))
        If TRIM$(VALACT1$(ARCHE$(Index))) <> "" Then
            COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
            COTEXT(Index).Refresh
            DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
            DETEXT(Index).Refresh
            DETEXT(Index).TabStop = False
        End If
        COTEXT(Index).SetFocus
    End If
End Sub
'
Private Sub COTEXT_LOSTFOCUS(Index As Integer)
    If TRIM$(COTEXT(Index).TEXT) <> "" Then
        TV11% = TIPFIELD(ARCHE$(Index), 1)
        LF11% = LENFIELD(ARCHE$(Index), 1)
        ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
        ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
        If ECO$ <> "" Then
             COTEXT(Index).TEXT = TRIM$(COTEXT(Index).TEXT)
             DETEXT(Index).TEXT = ECO$
             DETEXT(Index).Refresh
             DETEXT(Index).TabStop = False
           Else
             DETEXT(Index).TEXT = ""
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'
Private Sub COTEXT_GotFocus(Index As Integer)
    If NC% > (-1) Then Text1.SetFocus
End Sub
'

Private Sub COTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       If TRIM$(COTEXT(Index).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
               DETEXT(Index).TEXT = ECO$
               DETEXT(Index).Refresh
               DETEXT(Index).TabStop = False
             Else
               DETEXT(Index).TEXT = ""
               COTEXT(Index).SetFocus
          End If
       End If
    End If
End Sub

Private Sub DETEXT_CLICK(Index As Integer)
    COTEXT(Index).SetFocus
End Sub

Sub GENDARIBPRO()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARIBPRO")
    Call REPLA(REBLA$, USERNAME$, 41, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 63, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        If TICUEPRO%((-1) * NCLIX%) = 0 Then
          Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
          Call REPLA(RECORD$, Chr$(1), 3, 1)
          Call REPLA(RECORD$, MKS$(100), 5, 4)
          Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NCLIX%), RECORD$)
        End If
      End If
    Next U&
    Call CIERRARCH("DARIBPRO")
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command9_Click()
   NX% = FILEOPEN%(LUDAT$ + "PROVEEDO.DAT", 1, 128)
   While Not EOF(NX%)
   Line Input #NX%, AAX$
   NPX& = Val(Left$(AAX$, 6))
   If NPX& > 0 Then
     If NPX& <= 32767 Then
       Text1 = TRIM$(Str$(NPX&))
       NC% = (-1) * NPX&
       Text2 = TRIM$(Mid$(AAX$, 8, 41))
       Text3 = TRIM$(Mid$(AAX$, 8, 41))
       Text10 = TRIM$(Mid$(AAX$, 49, 16))
       Text11 = TRIM$(Mid$(AAX$, 97, 21))
       TEXT12 = TRIM$(LCase$(Mid$(AAX$, 118, 31)))
       Text4 = TRIM$(Mid$(AAX$, 149, 31))
       Text6 = TRIM$(Mid$(AAX$, 180, 31))
       COTEXT(0) = TRIM$(Mid$(AAX$, 211, 1))
       Text5 = TRIM$(Mid$(AAX$, 242, 8))
       Text15 = TRIM$(Mid$(AAX$, 251, 13))
       COTEXT(1) = "1"
       If VALICUIT%(Text15) < 1 Then COTEXT(1) = 0: Text15 = ""
       MODO% = 1
       Call GRADATPRO
     End If
   End If
   Wend
   Call CIERRARCH("*.*")
   Close #NX%
End Sub


