VERSION 5.00
Begin VB.Form NOCONF09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub-Sistema de Calidad"
   ClientHeight    =   9750
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame17 
      BackColor       =   &H00E0E0E0&
      Height          =   750
      Left            =   105
      TabIndex        =   89
      Top             =   8925
      Width           =   11145
      Begin VB.CommandButton Command35 
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
         Height          =   300
         Left            =   9870
         TabIndex        =   26
         Top             =   380
         Width           =   175
      End
      Begin VB.CommandButton Command34 
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
         Height          =   300
         Left            =   6930
         TabIndex        =   25
         Top             =   380
         Width           =   175
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
         Height          =   300
         Left            =   1470
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   " "
         Top             =   380
         Width           =   3330
      End
      Begin VB.Label Label13 
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
         Height          =   270
         Left            =   8360
         TabIndex        =   96
         Top             =   380
         Width           =   1485
      End
      Begin VB.Label Label12 
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
         Height          =   270
         Left            =   5400
         TabIndex        =   95
         Top             =   380
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Implementación"
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
         Index           =   26
         Left            =   8400
         TabIndex        =   94
         Top             =   105
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Notificado"
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
         Index           =   25
         Left            =   5460
         TabIndex        =   93
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre/Firma"
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
         Index           =   24
         Left            =   210
         TabIndex        =   92
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable de la No Conformidad"
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
         Index           =   23
         Left            =   210
         TabIndex        =   91
         Top             =   105
         Width           =   3795
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   105
      TabIndex        =   88
      Top             =   8295
      Width           =   11145
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
         Height          =   270
         Left            =   9555
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   1395
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
         Height          =   375
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "NOCONF09.frx":0000
         Top             =   150
         Width           =   8475
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N/C Nº"
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
         Index           =   22
         Left            =   8820
         TabIndex        =   90
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      Height          =   1065
      Left            =   105
      TabIndex        =   87
      Top             =   7245
      Width           =   11145
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
         Height          =   795
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "NOCONF09.frx":0002
         Top             =   210
         Width           =   10740
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   105
      TabIndex        =   84
      Top             =   6825
      Width           =   11145
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
         Height          =   300
         Left            =   1050
         TabIndex        =   20
         Top             =   105
         Width           =   175
      End
      Begin VB.Label Label18 
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
         Height          =   270
         Left            =   1260
         TabIndex        =   86
         Top             =   115
         Width           =   4110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Destino: "
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
         Index           =   21
         Left            =   210
         TabIndex        =   85
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Height          =   750
      Left            =   105
      TabIndex        =   82
      Top             =   6090
      Width           =   11145
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
         Height          =   375
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "NOCONF09.frx":0004
         Top             =   315
         Width           =   10785
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Acciones Correctivas (AC) y Preventivas (AP)  a Implementar  (indicar AC, AP, o AL)"
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
         Index           =   20
         Left            =   210
         TabIndex        =   83
         Top             =   105
         Width           =   9045
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   105
      TabIndex        =   73
      Top             =   5145
      Width           =   11145
      Begin VB.TextBox Text5 
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
         Height          =   330
         Index           =   3
         Left            =   7140
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox Text5 
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
         Height          =   330
         Index           =   2
         Left            =   7140
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox Text5 
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
         Height          =   330
         Index           =   1
         Left            =   1785
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox Text5 
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
         Height          =   330
         Index           =   0
         Left            =   1785
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   180
         Width           =   495
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
         Height          =   300
         Index           =   3
         Left            =   7665
         TabIndex        =   18
         Top             =   570
         Width           =   175
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
         Height          =   300
         Index           =   2
         Left            =   7665
         TabIndex        =   14
         Top             =   180
         Width           =   175
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
         Height          =   300
         Index           =   1
         Left            =   2310
         TabIndex        =   16
         Top             =   570
         Width           =   175
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
         Height          =   300
         Index           =   0
         Left            =   2310
         TabIndex        =   12
         Top             =   180
         Width           =   175
      End
      Begin VB.Label Label11 
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
         Height          =   330
         Index           =   3
         Left            =   7875
         TabIndex        =   81
         Top             =   570
         Width           =   3060
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Causas Ident."
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
         Index           =   19
         Left            =   5775
         TabIndex        =   80
         Top             =   675
         Width           =   1905
      End
      Begin VB.Label Label11 
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
         Height          =   330
         Index           =   2
         Left            =   7875
         TabIndex        =   79
         Top             =   180
         Width           =   3060
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Causas Ident."
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
         Index           =   18
         Left            =   5775
         TabIndex        =   78
         Top             =   285
         Width           =   1905
      End
      Begin VB.Label Label11 
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
         Height          =   330
         Index           =   1
         Left            =   2520
         TabIndex        =   77
         Top             =   570
         Width           =   2955
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Causas Ident."
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
         Index           =   17
         Left            =   420
         TabIndex        =   76
         Top             =   675
         Width           =   1380
      End
      Begin VB.Label Label11 
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
         Height          =   330
         Index           =   0
         Left            =   2520
         TabIndex        =   75
         Top             =   180
         Width           =   2955
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Causas Ident."
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
         Index           =   16
         Left            =   420
         TabIndex        =   74
         Top             =   285
         Width           =   1905
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   8400
      TabIndex        =   70
      Top             =   3360
      Width           =   2850
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1050
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   1545
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Index           =   15
         Left            =   105
         TabIndex        =   71
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   8400
      TabIndex        =   61
      Top             =   1680
      Width           =   2850
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
         Height          =   300
         Index           =   3
         Left            =   945
         TabIndex        =   5
         Top             =   1320
         Width           =   175
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
         Height          =   300
         Index           =   2
         Left            =   945
         TabIndex        =   4
         Top             =   960
         Width           =   175
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
         Height          =   300
         Index           =   1
         Left            =   945
         TabIndex        =   3
         Top             =   580
         Width           =   175
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
         Height          =   300
         Index           =   0
         Left            =   945
         TabIndex        =   2
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.F.:"
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
         Index           =   14
         Left            =   210
         TabIndex        =   69
         Top             =   1365
         Width           =   645
      End
      Begin VB.Label Label10 
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
         Height          =   330
         Index           =   3
         Left            =   1155
         TabIndex        =   68
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.F.:"
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
         Index           =   13
         Left            =   210
         TabIndex        =   67
         Top             =   1000
         Width           =   645
      End
      Begin VB.Label Label10 
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
         Height          =   330
         Index           =   2
         Left            =   1155
         TabIndex        =   66
         Top             =   950
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.F.:"
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
         Index           =   12
         Left            =   210
         TabIndex        =   65
         Top             =   650
         Width           =   645
      End
      Begin VB.Label Label10 
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
         Height          =   330
         Index           =   1
         Left            =   1155
         TabIndex        =   64
         Top             =   580
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.F.:"
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
         Index           =   11
         Left            =   210
         TabIndex        =   63
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label10 
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
         Height          =   330
         Index           =   0
         Left            =   1155
         TabIndex        =   62
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   105
      TabIndex        =   58
      Top             =   3360
      Width           =   8205
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
         Height          =   300
         Left            =   2625
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   1125
      End
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   3780
         TabIndex        =   9
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría"
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
         Index           =   10
         Left            =   1575
         TabIndex        =   60
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   3990
         TabIndex        =   59
         Top             =   210
         Width           =   4110
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   105
      TabIndex        =   55
      Top             =   2625
      Width           =   8205
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
         Height          =   300
         Left            =   3780
         TabIndex        =   7
         Top             =   210
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
         Height          =   300
         Left            =   2625
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Left            =   3990
         TabIndex        =   57
         Top             =   210
         Width           =   4110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre / Sector Detector "
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
         Index           =   9
         Left            =   210
         TabIndex        =   56
         Top             =   210
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   105
      TabIndex        =   53
      Top             =   4095
      Width           =   11145
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
         Height          =   585
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "NOCONF09.frx":0006
         Top             =   315
         Width           =   10720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable de la NO-Conformidad"
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
         Left            =   210
         TabIndex        =   54
         Top             =   105
         Width           =   3165
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   105
      TabIndex        =   51
      Top             =   1680
      Width           =   8205
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
         Height          =   300
         Left            =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   525
         Width           =   7950
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable de la NO Conformidad / acción Preventiva  (Proveedor Cliente Sector)"
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
         Left            =   210
         TabIndex        =   52
         Top             =   210
         Width           =   8100
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   8400
      TabIndex        =   47
      Top             =   945
      Width           =   2850
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1680
         TabIndex        =   50
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label8 
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
         Height          =   330
         Left            =   210
         TabIndex        =   49
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Antecedentes Nº y Fecha"
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
         Left            =   210
         TabIndex        =   48
         Top             =   105
         Width           =   2325
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4620
      TabIndex        =   44
      Top             =   945
      Width           =   3690
      Begin VB.CommandButton Command1 
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
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label7 
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
         Height          =   330
         Left            =   420
         TabIndex        =   46
         Top             =   315
         Width           =   3060
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código (Plano) / PSC"
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
         Left            =   945
         TabIndex        =   45
         Top             =   105
         Width           =   2010
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   105
      TabIndex        =   41
      Top             =   945
      Width           =   4425
      Begin VB.Label Label6 
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
         Height          =   330
         Left            =   210
         TabIndex        =   43
         Top             =   315
         Width           =   4005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación del Material / Proceso/Proced."
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
         Left            =   210
         TabIndex        =   42
         Top             =   105
         Width           =   4005
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   8400
      TabIndex        =   36
      Top             =   0
      Width           =   2850
      Begin VB.Label Label5 
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
         Height          =   330
         Left            =   1470
         TabIndex        =   40
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1050
         TabIndex        =   39
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   210
         TabIndex        =   38
         Top             =   525
         Width           =   840
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Nº:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   525
         TabIndex        =   37
         Top             =   105
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   10095
      Left            =   11445
      ScaleHeight     =   10035
      ScaleWidth      =   1320
      TabIndex        =   28
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command15 
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
         Picture         =   "NOCONF09.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Graba - Guardar e Imprime "
         Top             =   7980
         Width           =   645
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
         Left            =   120
         Picture         =   "NOCONF09.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "NOCONF09.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   960
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   95
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   30
         Top             =   9000
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "NOCONF09.frx":0766
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   9285
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   105
      TabIndex        =   33
      Top             =   0
      Width           =   8205
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CORRECTIVA Y PREVENTIVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1995
         TabIndex        =   35
         Top             =   525
         Width           =   4530
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DEMANDA DE ACCION "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2100
         TabIndex        =   34
         Top             =   105
         Width           =   4530
      End
   End
End
Attribute VB_Name = "NOCONF09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Command1.Enabled = False
    Call SELECHO("MASTER")
    CIX% = XVALO(VALACT1$("MASTER"))
    If CIX% < 1 Then GoTo 99
    '
    Label7 = VALACT1$("MASTER")
    Label6 = VALACT2$("MASTER")
99  Command1.Enabled = True

End Sub

Private Sub Command11_Click()
   Unload Me
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    Call SELECHO("TADESTI")
    TADE% = XVALO(VALACT1$("TADESTI"))
    If TADE% < 1 Then GoTo 99
    
    Label18 = VALACT2$("TADESTI")
99  Command14.Enabled = True
    '
End Sub

Private Sub Command15_Click()
       Command15.Enabled = False
       '
       HOII% = HOY(HOS$)
       FEMI% = FECHANUM(Label5)
       If FEMI% < 1 Or FEMI% > HOII% Then
          Call COMUNI("Fecha de Emisión no Válida")
          GoTo 99
       End If
       '
       'OBTENGO EL PROXIMO COMPROBANTE Y VALIDO QUE SEA EL MISMO
       '******************************
       CONDI$ = "NUMENC > 0"
10     NUMNC& = 1 + XVALO(VALOBADA("ACCORREC", "MAX(NUMENC)", CONDI$, "NOMESS"))
       '
       If NUMNC& <> XVALO(Label4) Then
          NOCONF$ = TRIM$(Str$(NUMNC&))
          Call MENSERR(24, "Número de Acción Correctiva '" + NOCONF$ + "' Ocupado.\   \FLEXOFT ha Re-Numerado La No Conformidad \como Número '" + NOCONF$ + "'.\  \Aceptar")
          Label4 = NOCONF$
          GoTo 10
       End If
       '
       'CREO EL REGISTRO
       '****************
       CONDI$ = "NumeNC = " & NUMNC&
       Call CreaRegistro("ACCORREC", "NumeNC", NUMNC&, "NOMESS")
       '
       Call ACTUREGISTRO("ACCORREC", "NumeNC", CONDI$, XVALO(Label4), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "FechaNC", CONDI$, CVDAT(FEMI%), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Cod_Inte", CONDI$, CODINT%(Label7), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Descripcion", CONDI$, TRIM$(Label6), REGAF&, "NOMESS")
          FANTCNCF% = FECHANUM(Label9)
       Call ACTUREGISTRO("ACCORREC", "FechaAntecNC", CONDI$, CVDAT(FANTCNCF%), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "AntecNC", CONDI$, TRIM$(Label8), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "ResponNC", CONDI$, Label5, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Catergoria", CONDI$, XVALO(Text3), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Canatidad", CONDI$, XVALO(Text4), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "IDSUBOR", CONDI$, TRIM$(Label10(0)), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "IDSUBOR1", CONDI$, TRIM$(Label10(1)), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "IDSUBOR2", CONDI$, TRIM$(Label10(2)), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "IDSUBOR3", CONDI$, TRIM$(Label10(3)), REGAF&, "NOMESS") '
    '   Call ACTUREGISTRO("ACCORREC", "ORDFAB", CONDI$, Label5, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "DescripcionNC", CONDI$, Text1, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "TipoCausa", CONDI$, Label11(0), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "TipoCausa1", CONDI$, Label11(1), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "TipoCausa2", CONDI$, Label11(2), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "TipoCausa3", CONDI$, Label11(3), REGAF&, "NOMESS")
'       Call ACTUREGISTRO("ACCORREC", "Causas", CONDI$, Label5, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "DescrpCausa", CONDI$, Text13, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Destino", CONDI$, Label5, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "DescripDestino", CONDI$, Text6, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "DescripAccCorr", CONDI$, Text8, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "ResponsableNC", CONDI$, TRIM$(Text10), REGAF&, "NOMESS")
           FNOTI% = FECHANUM(Label12)
       Call ACTUREGISTRO("ACCORREC", "FechaNotificacion", CONDI$, CVDAT(FNOTI%), REGAF&, "NOMESS")
           FIMP% = FECHANUM(Label13)
       Call ACTUREGISTRO("ACCORREC", "FechaImplementacion", CONDI$, CVDAT(FIMP%), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "Status", CONDI$, 1, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "CierreTec", CONDI$, Label5, REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "SecDetectoNC", CONDI$, TRIM$(Text2), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ACCORREC", "NotaCredito", CONDI$, TRIM$(Text9), REGAF&, "NOMESS")
       Call GRACONCATENAR
       '
       Call LIMPIAR_INFO
       CONDI$ = "NUMENC > 0"
       Label4 = 1 + XVALO(VALOBADA("ACCORREC", "MAX(NUMENC)", CONDI$, "NOMESS"))
       HOII% = HOY(HOS$)
       Label5 = FECHATEX$(HOII%)
99 Command15.Enabled = True

End Sub

Sub GRACONCATENAR()
        'GUARDO LAS ORDENES DE FABRICACION CONCATENADAS EN EL CAMPO ORDFAB
        '*****************************************************************
        NUMNC& = XVALO(Label4)
        ORFA$ = TRIM$(Label10(0))
        ORFA1$ = TRIM$(Label10(1))
        ORFA2$ = TRIM$(Label10(2))
        ORFA3$ = TRIM$(Label10(3))
        If ORFA$ <> "" Then CONCAORFA$ = ORFA$ + ";"
        If ORFA$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA1$ + ";"
        If ORFA2$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA2$ + ";"
        If ORFA3$ <> "" Then CONCAORFA$ = CONCAORFA$ + ORFA3$ + ";"
        If InStr(CONCAORFA$, ";") > 0 Then CONCAORFA$ = Left$(CONCAORFA$, Len(CONCAORFA$) - 1)
        '
        'GUARDO LAS CAUSAS CONCATENADAS EN EL CAMPO CAUSA
        '************************************************
        CAUSA$ = TRIM$(Text5(0))
        CAUSA1$ = TRIM$(Text5(1))
        CAUSA2$ = TRIM$(Text5(2))
        CAUSA3$ = TRIM$(Text5(3))
        If CAUSA$ <> "" Then CONCACAUSA$ = CAUSA$ + "-"
        If CAUSA1$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA1$ + "-"
        If CAUSA2$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA2$ + "-"
        If CAUSA3$ <> "" Then CONCACAUSA$ = CONCACAUSA$ + CAUSA3$ + "-"
        If InStr(CONCACAUSA$, "-") > 0 Then CONCACAUSA$ = Left$(CONCACAUSA$, Len(CONCACAUSA$) - 1)
        '
        CONDI$ = "NUMENC  = " & NUMNC&
        Call ACTUREGISTRO("ACCORREC", "ORDFAB", CONDI$, CONCAORFA$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ACCORREC", "Causas", CONDI$, CONCACAUSA$, REGAF&, "NOMESS")
End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("TACATEGO")
    TAC% = TRIM$(VALACT1$("TACATEGO"))
    If TAC% < 1 Then GoTo 99
    
    Text3 = TAC%
99  Command2.Enabled = True

End Sub

Private Sub Command3_Click(Index As Integer)

    Command3(Index).Enabled = False
    Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    
    Label10(Index) = VALACT1$("ECORDEF")
99  Command3(Index).Enabled = True

End Sub


Private Sub Command34_Click()
    Command34.Enabled = False
    Call SELECHO("SELFECHA")
    Label12 = VALACT1$("SELFECHA")
    Command34.Enabled = True
End Sub

Private Sub Command35_Click()
    Command35.Enabled = False
    Call SELECHO("SELFECHA")
    Label13 = VALACT1$("SELFECHA")
    Command35.Enabled = True

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call SELECHO("DEUDOR1")
    NC% = TRIM$(VALACT1$("DEUDOR1"))
    If NC% < 1 Then GoTo 99
    
    Text2 = VALACT1$("DEUDOR1")
    
99  Command4.Enabled = True

End Sub


Private Sub Label15_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Command8_Click(Index As Integer)

    Command8(Index).Enabled = False
    Call SELECHO("TAMOTI1")
    TAMO$ = TRIM$(VALACT1$("TAMOTI1"))
    If TAMO$ = "" Then GoTo 99
    '
    Text5(Index) = VALACT1$("TAMOTI1")
99  Command8(Index).Enabled = True



End Sub

Private Sub Label4_Change()
   '
   'VALIDO QUE EL LABLE POSEA UN VALOR NUMERICO
   '*******************************************
   NUMENOCO& = XVALO(Label4)
   If NUMENOCO& < 1 Then Call LIMPIAR_INFO
   '
   'VALIDO SI ES UNA NUEVA NO CONFORMIDADAD O EXISTENTE
   '***************************************************
   CONDI$ = " NumeNC = " & NUMENOCO&
   If CantRegistros&("ACCORREC", CONDI$, "NOMESS") < 1 Then
      Call LIMPIAR_INFO
      Exit Sub
   End If
   '
   Label5 = CVDAT(CVINT(VALOBADA("ACCORREC", "FechaNC", CONDI$, "NOMESS")))
   CI% = XVALO(VALOBADA("ACCORREC", "Cod_Inte", CONDI$, "NOMESS"))
   Label7 = CODEXT$(CI%)
   Label6 = DESCRIT0(CI%)
   Label9 = CVDAT(CVINT(VALOBADA("ACCORREC", "FechaAntecNC", CONDI$, "NOMESS")))
   Label8 = VALOBADA("ACCORREC", "AntecNC", CONDI$, "NOMESS")
   Label5 = VALOBADA("ACCORREC", "ResponNC", CONDI$, "NOMESS")
   Text3 = XVALO(VALOBADA("ACCORREC", "Catergoria", CONDI$, "NOMESS"))
   Text4 = XVALO(VALOBADA("ACCORREC", "Canatidad", CONDI$, "NOMESS"))
   Label10(0) = XVALO(VALOBADA("ACCORREC", "IDSUBOR", CONDI$, "NOMESS"))
   Label10(1) = XVALO(VALOBADA("ACCORREC", "IDSUBOR1", CONDI$, "NOMESS"))
   Label10(2) = XVALO(VALOBADA("ACCORREC", "IDSUBOR2", CONDI$, "NOMESS"))
   Label10(3) = XVALO(VALOBADA("ACCORREC", "IDSUBOR3", CONDI$, "NOMESS"))
   Text1 = VALOBADA("ACCORREC", "DescripcionNC", CONDI$, "NOMESS")
   Text5(0) = VALOBADA("ACCORREC", "TipoCausa", CONDI$, "NOMESS")
   Text5(1) = VALOBADA("ACCORREC", "TipoCausa1", CONDI$, "NOMESS")
   Text5(2) = VALOBADA("ACCORREC", "TipoCausa2", CONDI$, "NOMESS")
   Text5(3) = VALOBADA("ACCORREC", "TipoCausa3", CONDI$, "NOMESS")
   Text13 = VALOBADA("ACCORREC", "DescrpCausa", CONDI$, "NOMESS")
   Label5 = XVALO(VALOBADA("ACCORREC", "Destino", CONDI$, "NOMESS"))
   Text6 = VALOBADA("ACCORREC", "DescripDestino", CONDI$, "NOMESS")
   Text8 = VALOBADA("ACCORREC", "DescripAccCorr", CONDI$, "NOMESS")
   Text10 = VALOBADA("ACCORREC", "ResponsableNC", CONDI$, "NOMESS")
   Label12 = CVDAT(CVINT(VALOBADA("ACCORREC", "FechaNotificacion", CONDI$, "NOMESS")))
   Label13 = CVDAT(CVINT(VALOBADA("ACCORREC", "FechaImplementacion", CONDI$, "NOMESS")))
   Label5 = VALOBADA("ACCORREC", "CierreTec", CONDI$, "NOMESS")
   Text2 = VALOBADA("ACCORREC", "SecDetectoNC", CONDI$, "NOMESS")
   Text9 = VALOBADA("ACCORREC", "NotaCredito", CONDI$, "NOMESS")

    
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Text2_Change()
    NC% = XVALO(Text2)
    Label1 = RASOCLI$(NC%)
End Sub


Private Sub Text2_DblClick()
    Call Command4_Click
End Sub


Private Sub Text3_Change()
   CAT% = XVALO(Text3)
   Label2 = TRIM$(ECOARCH$("TACATEGO", Chr(CAT%)))
End Sub


Private Sub Text3_DblClick()
    Call Command2_Click
End Sub


Private Sub Text5_Change(Index As Integer)
    TIPCAU$ = TRIM$(Text5(Index))
    
    Label11(Index) = ECOARCH$("TAMOTI1", TIPCAU$)

End Sub

Private Sub Text5_DblClick(Index As Integer)
   Call Command8_Click(Index)
End Sub

Sub LIMPIAR_INFO()
   Label6 = ""
   Label7 = ""
   Label8 = ""
   Label9 = ""
   Text7 = ""
   For I% = 0 To 3
      Label10(I%) = ""
      Text5(I%) = ""
      Label11(I%) = ""
   Next I%
   Text2 = ""
   Text3 = ""
   Label2 = ""
   Label1 = ""
   Text4 = ""
   Text1 = ""
   Text13 = ""
   Label18 = ""
   Text6 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Label12 = ""
   Label13 = ""
   
End Sub
