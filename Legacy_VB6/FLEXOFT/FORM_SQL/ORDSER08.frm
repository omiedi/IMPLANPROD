VERSION 5.00
Begin VB.Form ORDESER08 
   Caption         =   "Ordenes de Servicio"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Informe"
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
      Left            =   7560
      TabIndex        =   94
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2580
         TabIndex        =   1
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   1640
         TabIndex        =   110
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label50 
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
         Height          =   300
         Left            =   1635
         TabIndex        =   104
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de PARTE:"
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
         TabIndex        =   96
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   840
         TabIndex        =   95
         Top             =   840
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2175
      Left            =   6480
      TabIndex        =   90
      Top             =   3240
      Width           =   4095
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
         Height          =   300
         Index           =   3
         Left            =   2880
         TabIndex        =   56
         Top             =   1680
         Width           =   180
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
         Height          =   300
         Index           =   2
         Left            =   2880
         TabIndex        =   45
         Top             =   1320
         Width           =   180
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
         Height          =   300
         Index           =   1
         Left            =   2880
         TabIndex        =   34
         Top             =   960
         Width           =   180
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
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   23
         Top             =   600
         Width           =   180
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
         Height          =   300
         Index           =   3
         Left            =   2040
         TabIndex        =   54
         Top             =   1680
         Width           =   180
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
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   43
         Top             =   1320
         Width           =   180
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
         Height          =   300
         Index           =   1
         Left            =   2040
         TabIndex        =   32
         Top             =   960
         Width           =   180
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
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Top             =   600
         Width           =   180
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
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   57
         Text            =   " "
         Top             =   1680
         Width           =   855
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
         Height          =   315
         Index           =   3
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   55
         Text            =   " "
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Text            =   " "
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Text            =   " "
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Text            =   " "
         Top             =   960
         Width           =   1935
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
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   44
         Text            =   " "
         Top             =   1320
         Width           =   495
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
         Height          =   315
         Index           =   1
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   33
         Text            =   " "
         Top             =   960
         Width           =   495
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
         Height          =   315
         Index           =   0
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   22
         Text            =   " "
         Top             =   600
         Width           =   495
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
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   46
         Text            =   " "
         Top             =   1320
         Width           =   855
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
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   35
         Text            =   " "
         Top             =   960
         Width           =   855
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
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Materiales"
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
         TabIndex        =   93
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Unidad"
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
         TabIndex        =   92
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Cant."
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
         Left            =   3480
         TabIndex        =   91
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   5640
      TabIndex        =   86
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox Text34 
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
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   13
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text34 
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
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Text            =   " "
         Top             =   840
         Width           =   1935
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
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   180
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
         Height          =   300
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   180
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
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   180
      End
      Begin VB.TextBox Text34 
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
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   5
         Text            =   " "
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Text            =   " "
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "Hs.Func."
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
         Left            =   3960
         TabIndex        =   89
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Nº Serie"
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
         TabIndex        =   88
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Modelo"
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
         TabIndex        =   87
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   79
      Top             =   1320
      Width           =   5415
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Reparaciones"
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
         Index           =   5
         Left            =   3600
         TabIndex        =   85
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Garantia"
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
         Left            =   2280
         TabIndex        =   84
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Puesta en Marcha      "
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
         Index           =   3
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Puesta en Marcha Otros"
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
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Service"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   81
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Repuestos"
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
         Index           =   0
         Left            =   3600
         TabIndex        =   80
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   8535
      Left            =   10680
      ScaleHeight     =   8475
      ScaleWidth      =   1155
      TabIndex        =   76
      ToolTipText     =   "Visualizar y Reimprimir"
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   710
         Left            =   75
         Picture         =   "ORDSER08.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   1920
         Width           =   710
      End
      Begin VB.CommandButton Command26 
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
         Height          =   710
         Left            =   75
         Picture         =   "ORDSER08.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   4680
         Width           =   710
      End
      Begin VB.CommandButton Command14 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   710
         Left            =   75
         Picture         =   "ORDSER08.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Aprueba, Graba e Imprime Orden de Servicio"
         Top             =   5520
         Width           =   710
      End
      Begin VB.CommandButton Command29 
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
         Height          =   720
         Left            =   75
         Picture         =   "ORDSER08.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1080
         Width           =   720
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
         Height          =   720
         Left            =   75
         Picture         =   "ORDSER08.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   240
         Width           =   720
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   6480
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -360
         Picture         =   "ORDSER08.frx":0EAA
         Top             =   6840
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Top             =   7920
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   66
      Top             =   3240
      Width           =   6255
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
         Index           =   3
         Left            =   480
         TabIndex        =   48
         Top             =   1680
         Width           =   180
      End
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
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   450
      End
      Begin VB.TextBox Text8 
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
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   49
         Text            =   " "
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text15 
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
         Height          =   315
         Index           =   3
         Left            =   3720
         TabIndex        =   50
         Text            =   " "
         Top             =   1680
         Width           =   615
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
         Height          =   315
         Index           =   3
         Left            =   4920
         TabIndex        =   51
         Text            =   " "
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text25 
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
         Height          =   315
         Index           =   3
         Left            =   5520
         TabIndex        =   52
         Text            =   " "
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text25 
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
         Height          =   315
         Index           =   0
         Left            =   5520
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text25 
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
         Height          =   315
         Index           =   1
         Left            =   5520
         TabIndex        =   30
         Text            =   " "
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text25 
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
         Height          =   315
         Index           =   2
         Left            =   5520
         TabIndex        =   41
         Text            =   " "
         Top             =   1320
         Width           =   615
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
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   615
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
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   29
         Text            =   " "
         Top             =   960
         Width           =   615
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
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   40
         Text            =   " "
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text15 
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
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   17
         Text            =   " "
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text15 
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
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   28
         Text            =   " "
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text15 
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
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   39
         Text            =   " "
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   27
         Text            =   " "
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         Height          =   315
         Index           =   2
         Left            =   3120
         TabIndex        =   38
         Text            =   " "
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   16
         Top             =   600
         Width           =   615
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
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   180
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
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   180
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
         Index           =   2
         Left            =   480
         TabIndex        =   37
         Top             =   1320
         Width           =   180
      End
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
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   450
      End
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   450
      End
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
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label51 
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
         Height          =   315
         Index           =   3
         Left            =   4320
         TabIndex        =   108
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label51 
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
         Height          =   315
         Index           =   2
         Left            =   4320
         TabIndex        =   107
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label51 
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
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   106
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label51 
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
         Height          =   315
         Index           =   0
         Left            =   4320
         TabIndex        =   105
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Permanencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   98
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Index           =   3
         Left            =   720
         TabIndex        =   97
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Hs. Viaje."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4920
         TabIndex        =   75
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "A:"
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
         Left            =   5520
         TabIndex        =   74
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "De:"
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
         Left            =   4920
         TabIndex        =   73
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "T. Hs."
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
         Left            =   4320
         TabIndex        =   72
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "A:"
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
         Left            =   3720
         TabIndex        =   71
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "De:"
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
         Left            =   3120
         TabIndex        =   70
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Personal"
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
         TabIndex        =   69
         Top             =   350
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Index           =   2
         Left            =   720
         TabIndex        =   99
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Index           =   0
         Left            =   720
         TabIndex        =   68
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Index           =   1
         Left            =   720
         TabIndex        =   67
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajo Efectuado"
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
      Left            =   120
      TabIndex        =   65
      Top             =   5520
      Width           =   10455
      Begin VB.TextBox Text53 
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
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   360
         Width           =   9975
      End
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
      Height          =   315
      Left            =   240
      MaxLength       =   5
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   855
   End
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
      Left            =   1080
      TabIndex        =   63
      Top             =   120
      Width           =   180
   End
   Begin VB.TextBox Text2 
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
      Height          =   300
      Left            =   1500
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   " "
      Top             =   120
      Width           =   4200
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   1500
      TabIndex        =   61
      TabStop         =   0   'False
      Text            =   " "
      Top             =   510
      Width           =   4200
   End
   Begin VB.TextBox Text4 
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
      Height          =   300
      Left            =   1500
      TabIndex        =   60
      TabStop         =   0   'False
      Text            =   " "
      Top             =   900
      Width           =   750
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
      Height          =   300
      Left            =   2340
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   " "
      Top             =   900
      Width           =   3360
   End
   Begin VB.Label Label26 
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
      Height          =   300
      Left            =   240
      TabIndex        =   64
      Top             =   510
      Width           =   1035
   End
End
Attribute VB_Name = "ORDESER08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE%))
   End If
End Sub

Private Sub Command10_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click(Index As Integer)
    Text43(Index).TEXT = ""
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    If COD$ <> "" Then
        Text43(Index).TEXT = COD$
    End If
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    If XVALO(Text1.TEXT) = 0 Then
        Call MENSERR(24, "Debe Seleccionar un Cliente"): GoTo 99
    End If
    '
    VUELTA% = 0
    For I& = 0 To 2
       If TRIM(Text26(I&)) <> "" Then
         If CODINT%(TRIM$(Text26(I&))) < 1 Then
            Call MENSERR(24, "Codigo de Modelo no Valido en Linea  " & I& + 1): GoTo 99
         End If
         If HORANUM(Text34(I&)) = 0 Then
            Call MENSERR(24, "Falta Horas de funcionamiento en linea en Linea  " & I& + 1): GoTo 99
         End If
       End If
    Next I&
    '
    For JJ% = 0 To 3
      If TRIM(Text43(JJ%)) <> "" Then
         If XVALO(Text35(JJ%)) <= 0 Then
            Call MENSERR(24, "Cantidad Invalida en Material de Linea " & Str(JJ% + 1)): GoTo 99
         End If
      End If
    Next JJ%
    '
    'GRABACION DE TABLA ENCAOSER
    SQLX$ = "SELECT * FROM ENCAOSER"
    Dim ENCASTMP As ADODB.Recordset
    Set ENCAOSMP = New ADODB.Recordset
    ENCAOSMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With ENCAOSMP
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       NUMORD& = XVALO(Label4)
       !NuOrdSe = NUMORD&
       !FeOrdSe = Label50
       !Nume_Cli = XVALO(Text1.TEXT)
       'OBTIENE EL TIPO DE ORDEN
       For I& = 0 To 5
          If Option1(I&).Value = True Then
             TIPORD& = I&: Exit For
          End If
       Next I&
       !TipoOrdSe = TIPORD&
       !Observa = Text53
       .Update
    End With
    ENCAOSMP.Close
    Set ENCAOSMP = Nothing
    '
    'GRABACION EN TABLA DETAOSER
    J% = 0
    SQLX$ = "SELECT * FROM DETAOSER"
    Dim DETAORTMP As ADODB.Recordset
    Set DETAORTMP = New ADODB.Recordset
    DETAORTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With DETAORTMP
      For U& = 1 To 3
        If TRIM(Text26(U& - 1)) <> "" Then ' Si no tiene codgio cargado no graba
          J% = J% + 1
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NuOrdSe = NUMORD&
          !NuOrden = J%
          CI% = CODINT(Text26(U& - 1))
          !Modelo = CI%
          !NumeroSerie = Text31(U& - 1)
          HH% = HORANUM(Text34(U& - 1))
          !HorasFunc = HH%
          .Update
        End If
      Next U&
    End With
    DETAORTMP.Close
    Set DETAORTMP = Nothing
    '
    Call APERBASDAT("STOCKS")
    CMOV$ = "VM" ' OS = ORDEN DE SERVICIO
    FECHA$ = FECHATEX(HOY(HO$))
    DIA$ = Mid$(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 2)
    DOSELA$ = CMOV$ & "." & ANIO$ & MES$ & DIA$
    DOSECO$ = DOSELA$
    REFE$ = "Consumo O/Servicio " & TRIM$(Label4)
    PROXNRO& = XVALO(ProNroComprobante(CMOV$))
    J1% = 0
    For JJ% = 0 To 3
      CI% = CODINT(Text43(JJ%))
      If CI% > 0 Then
         FF% = FECHANUM(Label50)
         DOPRILA$ = CMOV$ & "." & TRIM$(Str(PROXNRO&))
         DOPRICO$ = DOPRILA$
         J1% = J% + 1: NUORIT% = J1%
         NC% = XVALO(Text1)
         DEPO% = 1
         CANTI = (-1) * XVALO(Text35(0))
         Call MOVISTOK(FF%, CI%, "", CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, 0, 1, NC%, DEPO%, CANTI)
      End If
    Next JJ%
    '
    Call PRINTORSE
    '
    Call BLANCFORM
    NUMORD& = XVALO(VALOBADA("ENCAOSER", "MAX(NuOrdSe)", "NuOrdSe > 0")) + 1
    Label4 = NUMORD&
    '
99 Command14.Enabled = True

End Sub

Private Sub Command16_Click(Index As Integer)
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text38(Index).TEXT = VALONUE$
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   '
   Call SELECHO("PADRON")
   NLEG% = Val(TRIM$(VALACT1$("PADRON")))
   If NLEG% < 1 Then Exit Sub
   NC% = NLEG%
   Text6(Index).TEXT = TRIM$(Str$(NC%))
   Label2(Index).Caption = TRIM$(VALACT2$("PADRON"))
   '
End Sub



Private Sub Command26_Click()
   Command26.Enabled = False
   RTA1% = ALTERNA%("Configuración de Formulario \Padron de Operarios")
   If RTA1% < 1 Or RTA1% > 2 Then GoTo 99
   If RTA1% = 1 Then
     STPOSER08.Show 1
   Else
     Call FINAL("*PADRON")
     AppActivate Caption
     SendKeys Chr$(1)

   End If
   
99 Command26.Enabled = True
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    Call COPYSTRU("INDIENSA", "INDIORSE")
    Call SETULTREG("INDIORSE", 0)
    SQLX$ = "SELECT *  FROM ENCAOSER"
    '
    Dim ENCAOSTMP As ADODB.Recordset
    Set ENCAOSTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With ENCAOSTMP
          Do While Not .EOF
            NUMORD& = !NuOrdSe
            FECCHH% = CVINT(!FeOrdSe)
            NC% = !Nume_Cli
            RAZOSOC$ = RASOCLI$(NC%)
            '
            TTXX$ = REGBLAN$("INDIORSE")
            Call REPLA(TTXX$, MKI$(FECCHH%), 1, 2)
            Call REPLA(TTXX$, "Orden de Servicio " & NUMORD& & " - " & Left$(RAZOSOC$, 30), 3, 62)
            Call REGAPP$("INDIORSE", TTXX$)
            .MoveNext
          Loop
            '
        End With
     ENCAOSTMP.Close
     Set ENCAOSTMP = Nothing
     
     
10 If ULTREG&("INDIORSE") < 1 Then GoTo 99 ' SI TIENE INFORMACION MUESTRA EL INDICE
   Call SELECHO("INDIORSE")
   '
   NUDO$ = VALACT2$("INDIORSE") ' MUESTRO EL FORMULARIO IMPRESO  SELECCIONAD0
   If NUDO$ = "" Then GoTo 99
   AA = InStr(1, NUDO$, "-")
   NN = InStr(1, NUDO$, "-") - 19
   NUDOCU$ = TRIM$(Mid$(NUDO$, 18, NN))
   If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
   Screen.MousePointer = 11
   '
   TIDOCUM$ = "Orden de Servicio"
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        ARGBUS$ = TIDOCUM$ & " " & NUDOCU$ & " "
        If InStr(XX$, ARGBUS$) > 4 Then
           DOCUNU& = CVS(Left$(XX$, 4))
           GoTo 5
        End If
   Next UI&
   '
   Screen.MousePointer = 1
   Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Orden de Servicio.")
   Screen.MousePointer = 1
   GoTo 10
   '
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   Screen.MousePointer = 1
   GoTo 10
   '
99 Command4.Enabled = True
End Sub

Private Sub Command6_Click(Index As Integer)
Call Text26_DblClick(Index)
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Label50 = VALACT1$("SELFECHA")
   FF% = FECHANUM(Label50)
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "ORDSER08"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call APERBASDAT("STOCKS")
    Call CREATABLA_STOCKS
    'aSIGNO NUEMRO DE LA ORDEN DE SERVICIO
    NUMORD& = 1
    XX& = CantRegistros("ENCAOSER")
    If XX& > 1 Then NUMORD& = XVALO(VALOBADA("ENCAOSER", "MAX(NuOrdSe)", "NuOrdSe > 0")) + 1
    Label4 = NUMORD&
    '
    'CARGA FECHA AL FROMULARIO
    Label50 = FECHATEX$(HOY(HO$))
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub






Private Sub Text1_Change()
    '
    NC% = XVALO(Text1.TEXT)
    Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
    Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
    Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
    Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
    Label26 = CODICLI$(NC%)
    '
End Sub

Private Sub Text1_DblClick()
   Call Command1_Click
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   '
   If TRIM$(Text15(Index)) <> "" Then
    Text15(Index).TEXT = HORATEX$(HORANUM(Text15(Index).TEXT))
   End If
   '
   For I& = 0 To 3
     If TRIM$(Text8(I&)) <> "" And TRIM$(Text15(I&)) <> "" Then
       aaaa% = HORANUM(Text8(I&))
       bbbb% = HORANUM(Text15(I&))
       CCC = DIFHORA(aaaa%, bbbb%)
       difer$ = HORAMINU(60 * DIFHORA(aaaa%, bbbb%))
       'difer$ = HORATEX$(60 * 60 * DIFHORA(aaaa%, bbbb%))
       Label51(I&) = difer$
     End If
   Next I&

   
End Sub



'Private Sub Text19_LostFocus(INDEX As Integer)
'   If TRIM$(Text19(INDEX)) <> "" Then
'    Text19(INDEX).TEXT = HORATEX$(HORANUM(Text19(INDEX).TEXT))
'   End If
'End Sub

Private Sub Text22_LostFocus(Index As Integer)
   If TRIM$(Text22(Index)) <> "" Then
    Text22(Index).TEXT = HORATEX$(HORANUM(Text22(Index).TEXT))
   End If
End Sub

Private Sub Text25_LostFocus(Index As Integer)
   If TRIM$(Text25(Index)) <> "" Then
    Text25(Index).TEXT = HORATEX$(HORANUM(Text25(Index).TEXT))
   End If
End Sub

Private Sub Text26_DblClick(Index As Integer)
    Text26(Index).TEXT = ""
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    If COD$ <> "" Then
        Text26(Index).TEXT = COD$
    End If
End Sub


Private Sub Text34_LostFocus(Index As Integer)
    If TRIM$(Text34(Index)) <> "" Then
      Text34(Index).TEXT = HORATEX$(HORANUM(Text34(Index).TEXT))
    End If
End Sub

Private Sub Text35_LostFocus(Index As Integer)
    Text35(Index).TEXT = FORMATNUM$(XVALO(Text35(Index).TEXT), "F7.1")
End Sub

Private Sub Text43_DblClick(Index As Integer)
    Text43(Index).TEXT = ""
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    If COD$ <> "" Then
        Text43(Index).TEXT = COD$
        CIXI% = CODINT%(COD$)
        Text38(Index) = UNIMED$(CIXI%)
    End If
End Sub

Private Sub Text6_Change(Index As Integer)
    Label2(Index).Caption = ""
    Label2(Index).Caption = ECOARCH$("PADRON", MKI$(Val(Text6(Index).TEXT)))
End Sub

Private Sub Text6_DblClick(Index As Integer)
   Call SELECHO("PADRON")
   NLEG% = Val(TRIM$(VALACT1$("PADRON")))
   If NLEG% < 1 Then Exit Sub
   NC% = NLEG%
   Text6(Index).TEXT = TRIM$(Str$(NC%))
   Label2(Index).Caption = TRIM$(VALACT2$("PADRON"))
End Sub


Private Sub Text8_LostFocus(Index As Integer)
   If TRIM$(Text8(Index)) <> "" Then
    Text8(Index).TEXT = HORATEX$(HORANUM(Text8(Index).TEXT))
   End If
   For I& = 0 To 3
     
     If TRIM$(Text8(I&)) <> "" And TRIM$(Text15(I&)) <> "" Then
       aaaa% = HORANUM(Text8(I&))
       bbbb% = HORANUM(Text15(I&))
       CCC = DIFHORA(aaaa%, bbbb%)
       difer$ = HORAMINU(60 * DIFHORA(aaaa%, bbbb%))
       Label51(I&) = difer$
     End If
   Next I&
End Sub

Sub BLANCFORM()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text53 = ""
    Option1(3).Value = True
    For I& = 0 To 2
        Text26(I&) = ""
        Text31(I&) = ""
        Text34(I&) = ""
    Next I&
    For I& = 0 To 3
        Text6(I&) = ""
        Text8(I&) = ""
        Text15(I&) = ""
        Label51(I&) = ""
        Text22(I&) = ""
        Text25(I&) = ""
        Text43(I&) = ""
        Text38(I&) = ""
        Text35(I&) = ""
        Label2(I&) = ""
    Next I&
End Sub

Sub PRINTORSE()

    '
    TIPODOC$ = "Orden de Servicio" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Label4)
    '
    CODPARAM$(1) = "RASO-CLI": DOCPARAM$(1) = RASOCLI$(XVALO(Text1))  'RAZON SOCIAL CLIENTE
    CODPARAM$(2) = "DOMI-CLI": DOCPARAM$(2) = DOMICLI$(XVALO(Text1))  'DOMICILIO CLIENTE
    CODPARAM$(3) = "LOCA-CLI":  DOCPARAM$(3) = LOCACLI$(XVALO(Text1)) 'LOCALIDAD CLIENTE
    CODPARAM$(4) = "NUME-COM":  DOCPARAM$(4) = TRIM$(Label4)          'NUMERO DE DOCUMENTO
    CODPARAM$(5) = "FECH-EMI":  DOCPARAM$(5) = Label50             'FECHA SOLICITADA
    
    Select Case TIPORD&
       Case 0
         DESTI$ = "Repuestos"
       Case 1
         DESTI$ = "Service"
       Case 2
         DESTI$ = "Puesta en Marcha Otros"
       Case 3
         DESTI$ = "Puesta en Marcha"
       Case 4
         DESTI$ = "Garantia"
       Case 5
         DESTI$ = "Reparaciones"
     End Select
    CODPARAM$(6) = "DESTINO":  DOCPARAM$(6) = DESTI$
    CODPARAM$(7) = "REF-MAT3":  DOCPARAM$(7) = TRIM$(Text53.TEXT)
    CAPARDOC% = 7
    '
    CODTABU1$(1) = "DES-MPRI"
    CODTABU1$(2) = "UNI-MPRI"
    CODTABU1$(3) = "CAN-MPRI"
    CODTABU1$(4) = "DES-OPER"
    CODTABU1$(5) = "TIE-PREP"
    CODTABU1$(6) = "TIE-PROD"
    CODTABU1$(7) = "TIE-TOTA"
    CODTABU1$(8) = "TOT-PREP"
    CODTABU1$(9) = "TOT-PROD"
    CACOLTAB1% = 9
    '
    CAITAB1% = 0
    For U& = 1 To 3
       If TRIM$(Text43(U& - 1)) <> "" Then
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = TRIM$(Text43(U& - 1))  ' CODIGO DE MATERIAL
         TETABU1$(2, CAITAB1%) = TRIM$(Text38(U& - 1))  ' UNIDAD
         TETABU1$(3, CAITAB1%) = TRIM$(Text35(U& - 1))  ' CANTIDAD
         TETABU1$(4, CAITAB1%) = TRIM$(Label2(U& - 1)) ' PERSONAL
         TETABU1$(5, CAITAB1%) = TRIM$(Text8(U& - 1))  ' DESDE
         TETABU1$(6, CAITAB1%) = TRIM$(Text15(U& - 1))  ' HASTA
         TETABU1$(7, CAITAB1%) = TRIM$(Label51(U& - 1))  ' TOTAL
         TETABU1$(8, CAITAB1%) = TRIM$(Text22(U& - 1))  ' DESDE
         TETABU1$(9, CAITAB1%) = TRIM$(Text25(U& - 1))  ' HASTA
       End If
    Next U&
    '
    CODTABU2$(1) = "COD-MAT"
    CODTABU2$(2) = "REF-MAT1"
    CODTABU2$(3) = "REF-MAT2"
    CACOLTAB2% = 3
    '
    CAITAB2% = 0
    For U& = 1 To 3
       If TRIM$(Text26(U& - 1)) <> "" Then
         CAITAB2% = CAITAB2% + 1
         TETABU2$(1, CAITAB2%) = TRIM$(Text26(U& - 1))  ' CODIGO DE MATERIAL
         TETABU2$(2, CAITAB2%) = TRIM$(Text31(U& - 1))  ' NUMERO DE SERIE
         TETABU2$(3, CAITAB2%) = TRIM$(Text34(U& - 1))  ' HORAS DE FUNCIONAMIENTO
       End If
    Next U&
    '
    '
    FORIPRE$ = TRIM$(CONTROL$("", "ORDESERV"))
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub



