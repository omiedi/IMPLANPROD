VERSION 5.00
Begin VB.Form MODICLI 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Modificación de Datos Cuenta de Cliente"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   10530
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
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
      Height          =   285
      Left            =   1995
      TabIndex        =   70
      Text            =   " "
      Top             =   4095
      Width           =   3510
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Cta.Corr"
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
      Left            =   4830
      Picture         =   "MODICLI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Cliente"
      Top             =   6510
      Width           =   855
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
      Left            =   8715
      TabIndex        =   68
      Text            =   " "
      Top             =   4725
      Width           =   1710
   End
   Begin VB.TextBox Text7 
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
      Left            =   4440
      TabIndex        =   66
      Text            =   " "
      Top             =   4710
      Width           =   1815
   End
   Begin VB.ListBox LISORT 
      Height          =   1815
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   64
      Top             =   525
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton BOTSELCUE 
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
      Left            =   10245
      TabIndex        =   63
      Top             =   210
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
      Index           =   7
      Left            =   7665
      TabIndex        =   62
      Top             =   5985
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
      Index           =   6
      Left            =   10265
      TabIndex        =   61
      Top             =   5160
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
      Index           =   5
      Left            =   2520
      TabIndex        =   60
      Top             =   5985
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
      TabIndex        =   59
      Top             =   5565
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
      TabIndex        =   58
      Top             =   4740
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
      Left            =   6825
      TabIndex        =   57
      Top             =   3675
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
      Left            =   6825
      TabIndex        =   56
      Top             =   2835
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
      TabIndex        =   55
      Top             =   2310
      Width           =   175
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
      Left            =   5775
      Picture         =   "MODICLI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6510
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
      Left            =   2940
      Picture         =   "MODICLI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6510
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
      Top             =   4740
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
      Index           =   1
      Left            =   6510
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   3255
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
      Index           =   1
      Left            =   6510
      TabIndex        =   12
      Text            =   " "
      Top             =   2835
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
      Height          =   285
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
      Left            =   9030
      TabIndex        =   0
      Text            =   " "
      Top             =   210
      Width           =   1410
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
      Left            =   6720
      Picture         =   "MODICLI.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6510
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
      Picture         =   "MODICLI.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6510
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
      Picture         =   "MODICLI.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6510
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
      Left            =   8610
      Picture         =   "MODICLI.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6510
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
      Left            =   8085
      TabIndex        =   8
      Text            =   " "
      Top             =   2310
      Width           =   2355
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
      Height          =   285
      Left            =   1995
      TabIndex        =   9
      Text            =   " "
      Top             =   2835
      Width           =   3510
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
      Left            =   1995
      TabIndex        =   10
      Text            =   " "
      Top             =   3255
      Width           =   3510
   End
   Begin VB.TextBox TEXT12 
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
      Left            =   1995
      TabIndex        =   11
      Text            =   " "
      Top             =   3675
      Width           =   3510
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
      Left            =   8085
      TabIndex        =   14
      Top             =   2835
      Width           =   2355
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
      Left            =   6510
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   4095
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
      Index           =   2
      Left            =   6510
      TabIndex        =   15
      Text            =   " "
      Top             =   3675
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
      Left            =   8085
      TabIndex        =   17
      Text            =   " "
      Top             =   3675
      Width           =   2355
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
      Top             =   5565
      Width           =   570
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
      Left            =   1995
      TabIndex        =   22
      Text            =   " "
      Top             =   5985
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
      Left            =   1995
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5145
      Width           =   4245
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
      Top             =   5565
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
      Index           =   5
      Left            =   2835
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5985
      Width           =   3405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   230
      TabIndex        =   33
      Top             =   6510
      Width           =   2535
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   0
         Picture         =   "MODICLI.frx":1386
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8925
      End
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
      Index           =   6
      Left            =   8190
      TabIndex        =   24
      Text            =   " "
      Top             =   5145
      Width           =   2250
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
      Index           =   6
      Left            =   6510
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5565
      Width           =   3930
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
      Index           =   7
      Left            =   7980
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "  "
      Top             =   5985
      Width           =   2460
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
      Index           =   7
      Left            =   7140
      TabIndex        =   26
      Text            =   " "
      Top             =   5985
      Width           =   570
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto:"
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
      TabIndex        =   71
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Límite de Crédito:"
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
      Left            =   6720
      TabIndex        =   67
      Top             =   4830
      Width           =   1905
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.Std:"
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
      Left            =   2640
      TabIndex        =   65
      Top             =   4830
      Width           =   1710
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lista  Precios:"
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
      TabIndex        =   53
      Top             =   4845
      Width           =   1710
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "P-IVA"
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
      Left            =   5040
      TabIndex        =   52
      Top             =   2940
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
      TabIndex        =   51
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
      TabIndex        =   50
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
      Left            =   6510
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      Left            =   6090
      TabIndex        =   44
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
      TabIndex        =   43
      Top             =   2940
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
      TabIndex        =   42
      Top             =   3360
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
      TabIndex        =   41
      Top             =   3780
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
      Left            =   6090
      TabIndex        =   40
      Top             =   2940
      Width           =   1905
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cat I-B"
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
      Left            =   5040
      TabIndex        =   39
      Top             =   3780
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
      Left            =   6090
      TabIndex        =   38
      Top             =   3780
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
      TabIndex        =   37
      Top             =   5670
      Width           =   1710
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
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
      TabIndex        =   36
      Top             =   6090
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
      TabIndex        =   35
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
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
      Left            =   6300
      TabIndex        =   34
      Top             =   6090
      Width           =   765
   End
End
Attribute VB_Name = "MODICLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTSEL_Click(Index As Integer)
    If NC% < 1 Then
        Text1.SetFocus
        Exit Sub
      Else
        COTEXT(Index).TEXT = ""
        Call SELECHO(ARCHE$(Index))
        COTEXT(Index).SetFocus
    End If
End Sub

Private Sub BOTSELCUE_Click(Index As Integer)
    If NCLILLAMA% > 0 Then Exit Sub
    If MODO% = 2 Then
        Text1.TEXT = ""
        Call SELECHO("DEUDOR1")
        Text1.SetFocus
        Exit Sub
      Else
        Text1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Command1_Click()
    If NCLILLAMA% > 0 Then Exit Sub
    If NC% <> Val(Text1.TEXT) Then
       Call MENSERR(24, "Imposible Grabar - Intente de Nuevo")
       Call CIERRARCH("*.*")
       Call FINAL("")
       Exit Sub
    End If
    Call GRADATCLI
End Sub

Private Sub Command2_Click()
    '
    If NCLILLAMA% > 0 Then Exit Sub
    If MODO% = 1 Then
        Call CLEARFORMCLI
        NC% = 0
        Text1.TEXT = ""
        MODICLI.Refresh
        Text1.SetFocus
        Exit Sub
    End If
    '
    If NC% > 0 Then
        Call CARDATCLI(NC%)
        MODICLI.Refresh
      Else
        Text1.SetFocus
    End If
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Unload MODICLI
End Sub

Private Sub Command4_Click()
    Call HELPONLINE("AMACLI")
End Sub
'

Private Sub Command5_Click()
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% > 0 And ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
        Command5.Enabled = False
        Call NOTACLI(NCLIE%, 2)
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Cliente")
    End If
End Sub

Private Sub Command6_Click()
    If NC% > 0 Then
       PRINTFORM
    End If
End Sub
'

Private Sub Command9_Click()
   If NC% > 0 Then
      PROLLAMA$ = "MODICLI"
      RESCUDEU.Text1.TEXT = TRIM$(Str$(NC%))
      RESCUDEU.Show
      RESCUDEU.Refresh
   End If
End Sub

Private Sub Form_Activate()
    '
    Static MODANT%, YACTI%
    '
    If MODO% <> MODANT% Then
        YACTI% = 0
        MODANT% = MODO%
        If MODO% = 1 Then
             MODICLI.Caption = "Apertura de Nueva Cuenta Corriente de Cliente"
             MODICLI.BackColor = &HFFFFC0
          Else
             MODICLI.Caption = "Modificación de Datos de Cuenta de Cliente"
             MODICLI.BackColor = &HC0C0FF
        End If
        '
        EPAC$ = TRIM$(EMPREAC$)
        If EPAC$ <> "" Then
           Caption = Caption + "  -  " + EPAC$
        End If
        '
        MODICLI.Text1.TEXT = ""
        If MODO% = 1 Then
            MODICLI.Text1.TEXT = NROPRO& + 1
        End If
        If NCLILLAMA% > 0 Then
            MODICLI.Text1.TEXT = NCLILLAMA%
        End If
    End If
    '
    If YACTI% = 0 Then
       Text1.SetFocus
       YACTI% = 1
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

Private Sub Form_Unload(Cancel As Integer)
    '
    Call SORTCLI(1)
    '
    Screen.MousePointer = 11
    If ULTREG&("DEUDOR1") > 0 Then
        XXX$ = REGLEIDO$("DEUDOR1", 1)
        Call GRAREG("DEUDOR1", XXX$, 1)
        Call CIERRARCH("DEUDOR1")
    End If
    Screen.MousePointer = 1
    '
    Call FRESHECHO("DEUDOR1")
    FORMUSEL.Hide
    MODICLI.Hide
    If Left$(COMMAX$, 3) = "NEW" Then
        Call FINAL("")
    End If
    If NCLILLAMA% > 0 Then
        Call FINAL("")
    End If
    '
End Sub

Private Sub Text1_Click()
    If NCLILLAMA% > 0 Then Exit Sub
    NC% = 0
    Text1.SetFocus
End Sub

Private Sub Text1_DBLCLICK()
    If NCLILLAMA% > 0 Then Exit Sub
    If MODO% = 2 Then
        Text1.TEXT = ""
        Call SELECHO("DEUDOR1")
        Exit Sub
      Else
        Text1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus()
    '
    If MODO% = 1 Then
        Exit Sub
    End If
    '
    If Text1.TEXT = "" Then
        Text1.TEXT = VALACT1$("DEUDOR1")
        Text1.Refresh
        NC% = Val(Text1.TEXT)
        If NC% > 0 Then
            Call CARDATCLI(NC%)
            MODICLI.Refresh
            Text2.SetFocus
            Exit Sub
        End If
    End If
    '
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If NCLILLAMA% > 0 Then Exit Sub
    If KeyAscii = 13 Then
        VALICLI% = 0
        If Val(Text1.TEXT) < 1 Or Val(Text1.TEXT) > 32767 Then
            NC% = 0
            Text1.TEXT = ""
            Text1.SetFocus
            Exit Sub
        End If
        '
        NC% = Val(Text1.TEXT)
        '
        If MODO% = 1 Then
            If REGICLI&(NC%) > 0 Then
                Call MENSERR(24, "Cuenta Existente en la Base de Datos")
                NC% = 0
                Text1.TEXT = ""
                If MODO% = 1 Then
                    Text1.TEXT = NROPRO& + 1
                End If
                Text1.SetFocus
                Exit Sub
              Else
                Text2.SetFocus
                Exit Sub
            End If
        End If
        '
        If REGICLI&(NC%) < 1 Then
            Call MENSERR(24, "La Cuenta" + Str$(NC%) + " Invocada no Figura\en la Base de Datos de CLIENTES")
            NC% = 0
            Text1.SetFocus
            Exit Sub
        End If
        '
        If NC% > 0 Then
              Call CARDATCLI(NC%)
              MODICLI.Refresh
              Text2.SetFocus
              Exit Sub
           Else
              Text1.SetFocus
        End If
    End If
End Sub
'
Private Sub Text1_LostFocus()
    '
    If MODO% = 1 Then
        '
        NC% = Val(Text1.TEXT)
        If REGICLI&(NC%) > 0 Then
             Call MENSERR(24, "Cuenta Existente en la Base de Datos")
             NC% = 0
             Text1.TEXT = ""
             If MODO% = 1 Then
                Text1.TEXT = NROPRO& + 1
             End If
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

Private Sub Text2_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text3_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text4_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text5_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text6_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text7_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text9_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text10_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub TexT11_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text12_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text13_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub Text15_GotFocus()
    If NC% < 1 Then Text1.SetFocus
End Sub
'

Private Sub COTEXT_DblClick(Index As Integer)
    If NC% < 1 Then
        Text1.SetFocus
        Exit Sub
      Else
        COTEXT(Index).TEXT = ""
        Call SELECHO(ARCHE$(Index))
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
             COTEXT(Index).SetFocus
             Exit Sub
        End If
      Else
        DETEXT(Index).TEXT = ""
    End If
End Sub
'
Private Sub COTEXT_GotFocus(Index As Integer)
    If NC% < 1 Then Text1.SetFocus
    If COTEXT(Index).TEXT = "" Then
        COTEXT(Index).TEXT = VALACT1$(ARCHE$(Index))
        COTEXT(Index).Refresh
        DETEXT(Index).TEXT = VALACT2$(ARCHE$(Index))
        DETEXT(Index).Refresh
        DETEXT(Index).TabStop = False
    End If
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

