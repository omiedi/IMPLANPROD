VERSION 5.00
Begin VB.Form COPUTIE07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Puesta a Tierra"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   645
      Left            =   10120
      Picture         =   "COPUTIE07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Graba e Imprime Informe"
      Top             =   9480
      Width           =   750
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1575
      TabIndex        =   4
      Text            =   " "
      Top             =   2160
      Width           =   2520
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
      Height          =   285
      Left            =   9690
      TabIndex        =   3
      Top             =   165
      Width           =   165
   End
   Begin VB.TextBox Text11 
      Height          =   300
      Left            =   8760
      TabIndex        =   2
      Text            =   " "
      Top             =   165
      Width           =   945
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00006000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8325
      Left            =   10120
      ScaleHeight     =   8265
      ScaleWidth      =   870
      TabIndex        =   37
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command7 
         Height          =   650
         Left            =   105
         Picture         =   "COPUTIE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Tabla de Instrumentos"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   105
         Picture         =   "COPUTIE07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprime Informe de Puesta a Tierra"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Height          =   650
         Left            =   105
         Picture         =   "COPUTIE07.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   38
         Top             =   7680
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Height          =   630
         Left            =   105
         Picture         =   "COPUTIE07.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "COPUTIE07.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   705
   End
   Begin VB.TextBox Text9 
      Height          =   765
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   8370
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8760
      TabIndex        =   1
      Top             =   720
      Width           =   165
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Continuidad de Puesta a Tierra "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   9735
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   9
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   " "
         Top             =   5820
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   9
         Left            =   4680
         TabIndex        =   94
         Text            =   " "
         Top             =   5820
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   9
         Left            =   3480
         TabIndex        =   93
         Text            =   " "
         Top             =   5820
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   8
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   " "
         Top             =   5290
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   8
         Left            =   4680
         TabIndex        =   88
         Text            =   " "
         Top             =   5290
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   8
         Left            =   3480
         TabIndex        =   87
         Text            =   " "
         Top             =   5290
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   7
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   " "
         Top             =   4750
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   7
         Left            =   4680
         TabIndex        =   82
         Text            =   " "
         Top             =   4750
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   7
         Left            =   3480
         TabIndex        =   81
         Text            =   " "
         Top             =   4750
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   6
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   " "
         Top             =   4210
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   6
         Left            =   4680
         TabIndex        =   76
         Text            =   " "
         Top             =   4210
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   6
         Left            =   3480
         TabIndex        =   75
         Text            =   " "
         Top             =   4210
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   5
         Left            =   3480
         TabIndex        =   68
         Text            =   " "
         Top             =   3680
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   5
         Left            =   4680
         TabIndex        =   67
         Text            =   " "
         Top             =   3680
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   5
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   " "
         Top             =   3680
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   4
         Left            =   3480
         TabIndex        =   62
         Text            =   " "
         Top             =   3160
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   4
         Left            =   4680
         TabIndex        =   61
         Text            =   " "
         Top             =   3160
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   4
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   " "
         Top             =   3160
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   3
         Left            =   3480
         TabIndex        =   56
         Text            =   " "
         Top             =   2640
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   3
         Left            =   4680
         TabIndex        =   55
         Text            =   " "
         Top             =   2640
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   3
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   " "
         Top             =   2640
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   2
         Left            =   3480
         TabIndex        =   50
         Text            =   " "
         Top             =   2120
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   2
         Left            =   4680
         TabIndex        =   49
         Text            =   " "
         Top             =   2120
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   2
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   " "
         Top             =   2120
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   1
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   1600
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   1
         Left            =   4680
         TabIndex        =   10
         Text            =   " "
         Top             =   1600
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   1
         Left            =   3480
         TabIndex        =   9
         Text            =   " "
         Top             =   1600
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   0
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   " "
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   0
         Left            =   4680
         TabIndex        =   7
         Text            =   " "
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   500
         Index           =   0
         Left            =   3480
         TabIndex        =   6
         Text            =   " "
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   9
         Left            =   7080
         TabIndex        =   98
         Top             =   5820
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   9
         Left            =   5880
         TabIndex        =   97
         Top             =   5820
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(10)"
         Height          =   500
         Index           =   8
         Left            =   240
         TabIndex        =   96
         Top             =   5820
         Width           =   3255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   8
         Left            =   7080
         TabIndex        =   92
         Top             =   5290
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   8
         Left            =   5880
         TabIndex        =   91
         Top             =   5290
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(9)"
         Height          =   500
         Index           =   7
         Left            =   240
         TabIndex        =   90
         Top             =   5290
         Width           =   3255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   7
         Left            =   7080
         TabIndex        =   86
         Top             =   4750
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   7
         Left            =   5880
         TabIndex        =   85
         Top             =   4750
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(8)"
         Height          =   500
         Index           =   6
         Left            =   240
         TabIndex        =   84
         Top             =   4750
         Width           =   3255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   6
         Left            =   7080
         TabIndex        =   80
         Top             =   4210
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   6
         Left            =   5880
         TabIndex        =   79
         Top             =   4210
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(7)"
         Height          =   500
         Index           =   5
         Left            =   240
         TabIndex        =   78
         Top             =   4210
         Width           =   3255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(6)"
         Height          =   500
         Index           =   4
         Left            =   240
         TabIndex        =   71
         Top             =   3680
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   5
         Left            =   5880
         TabIndex        =   70
         Top             =   3680
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   5
         Left            =   7080
         TabIndex        =   69
         Top             =   3680
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(5)"
         Height          =   500
         Index           =   3
         Left            =   240
         TabIndex        =   65
         Top             =   3160
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   4
         Left            =   5880
         TabIndex        =   64
         Top             =   3160
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   4
         Left            =   7080
         TabIndex        =   63
         Top             =   3160
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(2)"
         Height          =   500
         Index           =   2
         Left            =   240
         TabIndex        =   59
         Top             =   1600
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   3
         Left            =   5880
         TabIndex        =   58
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   3
         Left            =   7080
         TabIndex        =   57
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(3)"
         Height          =   500
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   2120
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   2
         Left            =   5880
         TabIndex        =   52
         Top             =   2120
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   7080
         TabIndex        =   51
         Top             =   2120
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   7080
         TabIndex        =   42
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   1
         Left            =   5880
         TabIndex        =   41
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   7080
         TabIndex        =   36
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   500
         Index           =   0
         Left            =   5880
         TabIndex        =   35
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(4)"
         Height          =   500
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(1)"
         Height          =   500
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   3250
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                   Veredicto"
         Height          =   690
         Left            =   8280
         TabIndex        =   32
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resistencia máxima per- mitida (Ohm)"
         Height          =   690
         Left            =   7080
         TabIndex        =   31
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resistencia Calculada (Ohm)"
         Height          =   690
         Left            =   5880
         TabIndex        =   30
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caida de tensión medida (V)"
         Height          =   690
         Left            =   4680
         TabIndex        =   29
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corriente de ensayo aplicada (A)"
         Height          =   690
         Left            =   3490
         TabIndex        =   28
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                               Ensayo aplicado entre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   3250
      End
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   7200
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label98 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7200
      TabIndex        =   73
      ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
      Top             =   1140
      Width           =   2640
   End
   Begin VB.Label Label97 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Instrumento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6090
      TabIndex        =   72
      Top             =   1185
      Width           =   1050
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   46
      Top             =   780
      Width           =   300
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   9480
      TabIndex        =   45
      ToolTipText     =   "Número de Revisión"
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Técnico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   44
      Top             =   2190
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   840
      Picture         =   "COPUTIE07.frx":1354
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "COPUTIE07.frx":165E
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      Top             =   1800
      Width           =   2520
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7905
      TabIndex        =   40
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veredicto General:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      TabIndex        =   27
      Top             =   3405
      Width           =   2220
   End
   Begin VB.Label Label100 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analizador de continuidad de puesta a tierra"
      Height          =   405
      Left            =   5880
      TabIndex        =   26
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label Label99 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   25
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label92 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   24
      Top             =   780
      Width           =   1380
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   23
      Top             =   1830
      Width           =   1380
   End
   Begin VB.Label Label88 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   525
      Left            =   480
      TabIndex        =   22
      Top             =   1200
      Width           =   3585
   End
   Begin VB.Label Label96 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   21
      Top             =   885
      Width           =   780
   End
   Begin VB.Label Label95 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   840
      Width           =   2865
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Informe de Puesta a Tierra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "COPUTIE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_COPUTIE07$
Private Sub Command1_Click()
   '\\\OT-08-0935-OD-08/10/08///
   Call CIERRARCH("*.*")
   Unload Me
   '\\\OT-08-0935-OD-FIN///
End Sub

Private Sub Command2_Click()
   '\\\OT-08-0935-OD-08/10/08///
    Command2.Enabled = False
    '
    'VALIDACION POR FECHA
    
    VDEVV$ = TRIM$(Text11.TEXT)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    'VALIDACION DE NUMERO DE SERIE
    '
    NUSERIE$ = TRIM$(Text8)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    If TRIM$(Label98) = "" Then
      Call MENSERR(24, "Falta Informar Instrumento")
      GoTo 99
    End If
    
    'VALIDACION DE VEREDICTO GENERAL
    '
    If TRIM$(Text7) = "" Then
      Call MENSERR(24, "Falta Informar Veredicto General")
      GoTo 99
    End If
    If TRIM$(Label11) = "" Then
      Call MENSERR(24, "Falta Informar Inspector")
      GoTo 99
    End If
    '
    If TRIM$(Text2) = "" Then
      Call MENSERR(24, "Falta Informar Técnico")
      GoTo 99
    End If
'    If UCase(TRIM$(Text7)) <> "A" Then
'       If UCase(TRIM$(Text5(0))) = "A" And UCase(TRIM$(Text5(1))) = "A" Then
'         Call MENSERR(24, "Si Ambos Veredictos Estan Aprobados Veredicto General Debe Ser 'A'")
'         GoTo 99
'       End If
'    End If
    '
    If UCase(TRIM$(Text7)) <> "A" Then
         APRO% = 0
         For i& = 0 To Text5.UBound
          If UCase(TRIM$(Text5(i&))) = "A" Then APRO% = APRO% + 1
         Next i&
       
         If APRO% = 10 Then
            Call MENSERR(24, "Si Todos los Veredictos Estan Aprobados Veredicto General Debe Ser 'A'")
            GoTo 99
         End If
    End If
    '
    
    CIXI% = CODINT%(Label95)
    Screen.MousePointer = 11
    '
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 5 = " & Now & vbCrLf
    
'    SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR070 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'    SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
''    Dim ENSA070TMP As ADODB.Recordset
''    Set ENSA070TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'    Set ENSA070TMP = READSET(SQLZ$)
'    ULTREVIS% = XVALO(ENSA070TMP!MAXREVI)
'    ENSA070TMP.Close
'    Set ENSA070TMP = Nothing
'TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 6 = " & Now & vbCrLf
'    SQLX$ = "SELECT REVISION FROM ENSAR070 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- VALIDAR EL  VALOR DE CAMPO REVISION
'    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
'    '
'TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 7 = " & Now & vbCrLf
''    Dim ENSA070TMP1 As ADODB.Recordset
''    Set ENSA070TMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    Set ENSA070TMP1 = READSET(SQLX$)
'    With ENSA070TMP1
'     If Not .EOF Then
'       REVIS% = 1 + ULTREVIS%
'     Else
'       REVIS% = 0 'SI ES EL PRIMERO EL VALOR VA A SER 0
'     End If
'    End With
'    ENSA070TMP1.Close
'    Set ENSA070TMP1 = Nothing

    CONDI$ = "NumeroSerie ='" & NUSERIE$ & "'"
    CANTRE& = CantRegistros&("ENSAR070", CONDI$, "NOMESS")
    If CANTRE& = 0 Then
       REVIS% = 0
    Else
       REVIS% = CANTRE& 'NO LE SUMO UNO POR QUE ARRANCA DE 0 (EJ. SI DEVUELVE 0 REGISTROS VA A DAR DE ALTA EL NUMERO 0, SI DEVUELVE 1 REGISTRO
                        'VA A DAR DE ALTA EL 1, POR QUE EL REGISTRO 1 CORRESPONDE AL 0
    End If

TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 8 = " & Now & vbCrLf
    '
    '\\\OT-08-0940-OD-20/10/08///
    SQLX$ = "SELECT * FROM ENSAR070"
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NUSERIE$ & "'"
    SQLX$ = SQLX$ + " AND REVISION =" & REVIS%
    Dim ENSATMP As ADODB.Recordset
    Set ENSATMP = New ADODB.Recordset
    ENSATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 9 = " & Now & vbCrLf
    
    With ENSATMP
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !NUMEROSERIE = TRIM$(Text8)
            !MODELO = CIXI%
            !Inspector = Left$(TRIM$(Label11), 32)
            !Comentarios = Text9
            !INSTRUMENTO = Left$(Label98, 32)
            '
            !Corr_Ensa1 = XVALO(Text1(0))
            !CTension1 = XVALO(Text3(0))
            !Veredicto1 = Left$(TRIM$(Text5(0)), 6)
            '
            !Corr_Ensa2 = XVALO(Text1(1))
            !CTension2 = XVALO(Text3(1))
            !Veredicto2 = Left$(TRIM$(Text5(1)), 6)
            '
            !Corr_Ensa3 = XVALO(Text1(2))
            !CTension3 = XVALO(Text3(2))
            !Veredicto3 = Left$(TRIM$(Text5(2)), 6)
            '
            !Corr_Ensa4 = XVALO(Text1(3))
            !CTension4 = XVALO(Text3(3))
            !Veredicto4 = Left$(TRIM$(Text5(3)), 6)
            '
            !Corr_Ensa5 = XVALO(Text1(4))
            !CTension5 = XVALO(Text3(4))
            !Veredicto5 = Left$(TRIM$(Text5(4)), 6)
            '
            !Corr_Ensa6 = XVALO(Text1(5))
            !CTension6 = XVALO(Text3(5))
            !Veredicto6 = Left$(TRIM$(Text5(5)), 6)
            '
            !Corr_Ensa7 = XVALO(Text1(6))
            !CTension7 = XVALO(Text3(6))
            !Veredicto7 = Left$(TRIM$(Text5(6)), 6)
            '
            !Corr_Ensa8 = XVALO(Text1(7))
            !CTension8 = XVALO(Text3(7))
            !Veredicto8 = Left$(TRIM$(Text5(7)), 6)
            '
            !Corr_Ensa9 = XVALO(Text1(8))
            !CTension9 = XVALO(Text3(8))
            !Veredicto9 = Left$(TRIM$(Text5(8)), 6)
            '
            !Corr_Ensa10 = XVALO(Text1(9))
            !CTension10 = XVALO(Text3(9))
            !Veredicto10 = Left$(TRIM$(Text5(9)), 6)
            '
            !VerGeneral = Left$(TRIM$(Text7), 6)
            !Bloqueado = "SI"
            !FECH_ENSA = Text11
            !TECNICO = Left$(TRIM$(Text2), 32) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-AGREGAR GRABACION DEL TECNICO
            !Revision = REVIS% ' GRABA CAMPO DE REVISION
            !INSTRUMENTO = Left$(Label98, 32)
            
            .Update
    End With
    '
    ENSATMP.Close
    Set ENSATMP = Nothing
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 10 = " & Now & vbCrLf
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Conformidad Puesta a Tierra" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Text2))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text7))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 11 = " & Now & vbCrLf
    'Call PRINTMEDI(1)
    
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(REVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("ENPT-MDT", FILTX$)
    '
        
    ULTREVIS% = XVALO(Label13)
    CIXI% = CODINT%(TRIM$(Label95))
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(ULTREVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("ENPT-MDT", FILTX$)
    
    
    Image1.Visible = True
    Image3.Visible = False
    Command23.Enabled = True
    '
    Call BLANFORMU
    '
    Screen.MousePointer = 1
    '
    '\\\OT-08-0940-OD-fin///
''GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TXTRAZA_COPUTIE07-" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_COPUTIE07$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%

99  Command2.Enabled = True

   '\\\OT-08-0935-OD-fin///

End Sub


Private Sub Command23_Click()
    NUSERIE$ = TRIM$(Text8)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       Exit Sub
    End If
    
    ULTREVIS% = XVALO(Label13)
    CIXI% = CODINT%(TRIM$(Label95))
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(ULTREVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("ENPT-MDT", FILTX$)

'    Call PRINTMEDI(0)
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPINFE.Show 1
   Command26.Enabled = True

End Sub

Private Sub Command4_Click()
Call Text8_DblClick
End Sub


Private Sub Command6_Click()
   '\\\OT-08-0935-OD-08/10/08///
    If Image3.Visible = False Then
      Call CommSel("SELFECHA", Command6, Text11)
    Else
      Command4.SetFocus
    End If
   '\\\OT-08-0935-OD-FIN///
End Sub


Private Sub Command7_Click()
   Command7.Enabled = False
   Call COFUTIE07.Command7_Click
   Command7.Enabled = True

End Sub

Private Sub Form_Load()
   Call ABREORFAB
  Call BLANFORMU
  FF% = HOY(HO$)
   Text11.TEXT = HO$
   Label11 = USERNAME$
End Sub

Private Sub Label98_Click()
   Call SELECHO("INSTRUME")
   CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
   If CODINSTRU% < 1 Then Exit Sub
   Label98 = ECOARCH$("INSTRUME", Chr$(CODINSTRU%))
End Sub

Private Sub Image1_DblClick()
   If DERACCE("MODENS70", "Modificacion de Informe de Ensayo") > 1 Then
        Image3.Visible = True
        Image1.Visible = False
        Command2.Enabled = False
        Command23.Enabled = True '\\\OT-08-1064-OD-21/11/08-MEDITEA- HABILITA LA IMPRESORA

   End If
End Sub

Private Sub Image3_DblClick()
   If DERACCE("MODENS70", "Modificacion de Informe de Ensayo") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False '\\\OT-08-1064-OD-21/11/08-MEDITEA- DESHABILITA LA IMPRESORA
        'PONGO FECHA DEL DIA CUANDO VA A REALIZAR OTRA REVISION
        HOII% = HOY(HOS$)
        Text11 = HOS$

   End If
End Sub

Private Sub Label9_Change(Index As Integer)
   If XVALO(Text1(Index)) > 0 Then
     If XVALO(Label9(Index)) <= XVALO(Label10(Index)) And XVALO(Label9(Index)) > 0 Then
       Text5(Index) = "A"
     Else
       Text5(Index) = "NC"
     End If
   Else
       Text5(Index) = ""
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   '\\\OT-08-0935-OD-08/10/08///
    If XVALO(Text1(Index)) > 0 And XVALO(Text1(Index)) > 0 Then
      Label9(Index) = XVALO(Text3(Index)) / XVALO(Text1(Index))
      Label9(Index) = FORMATNUM(Label9(Index), "F7.4")
    Else
      Label9(Index) = ""
    End If
    Call Text5_Change(Index)
    '\\\OT-08-0935-OD-FIN///
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
      Text3(Index).SetFocus
  End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  '\\\OT-08-0935-OD-08/10/08///
  VALOR1 = ROUND#(XVALO(Text1(Index)), 2)
  
  If XVALO(VALOR1) < 1000 Then
   Text1(Index) = FORMATNUM(Text1(Index), "F6.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 1000.\Debe Corregir Para Poder Continuar.")
    Text1(Index).SetFocus
  End If
  
  '\\\OT-08-0935-OD-FIN///
End Sub

Private Sub Text11_DblClick()
   '\\\OT-08-0935-OD-08/10/08///
   Call Command6_Click
   '\\\OT-08-0935-OD-FIN///
End Sub

Private Sub Text11_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text2_GotFocus()
  If Image3.Visible = True Then Command4.SetFocus '\\\OT-08-1066-OD-21/11/08-MEDITEA QUE NO SE PUEDA EDITAR SI TIENE EL CANDADO VISIBLE
End Sub


Private Sub Text3_Change(Index As Integer)
   '\\\OT-08-0935-OD-08/10/08///
   If XVALO(Text1(Index)) > 0 And XVALO(Text1(Index)) > 0 Then
     Label9(Index) = XVALO(Text3(Index)) / XVALO(Text1(Index))
     Label9(Index) = FORMATNUM(Label9(Index), "F7.4")
   Else
     Label9(Index) = ""
   End If
   Call Text5_Change(Index)
   '\\\OT-08-0935-OD-fin///
End Sub

Private Sub Text3_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index < 5 Then
      Text1(Index + 1).SetFocus
    Else
      Text1(0).SetFocus
    End If
  End If
End Sub

Private Sub Text3_LostFocus(Index As Integer)
  '\\\OT-08-0935-OD-08/10/08///
  
  VALOR1 = ROUND#(XVALO(Text3(Index)), 2)
  
  If XVALO(VALOR1) < 1000 Then
   Text3(Index) = FORMATNUM(Text3(Index), "F6.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 1000.\Debe Corregir Para Poder Continuar.")
    Text3(Index).SetFocus
  End If
  '\\\OT-08-0935-OD-FIN///

End Sub

Private Sub Text5_Change(Index As Integer)
   '
   Text7 = "" '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDACION PARA PONER A O NC EN VEREDICTO GENERAL
   '
   
   For i& = 0 To Text5.UBound
      If Text5(i&) = "NC" Then Text7 = "NC": Exit Sub
   Next i&
   '
   APRO% = 0
   For i& = 0 To Text5.UBound
      If UCase(TRIM$(Text5(i&))) = "A" Then APRO% = APRO% + 1
   Next i&
   If APRO% = 10 Then Text7 = "A": Exit Sub
   
   '
   'VALIDACION QUE SI LOS DATOS QUE ESTAN COMPLETOS ESTAN APROBADOS Y EL RESTO ESTA VACIO
   'EL VEREDICTO GENERAL SE CONSIDERA A (APROBADO)
   VACIO% = 0
   For i& = 0 To Text5.UBound
      If TRIM$(Text1(i&)) = "" And TRIM$(Text3(i&)) = "" And TRIM$(Text5(i&)) = "" Then VACIO% = VACIO% + 1
   Next i&
   
   If APRO% > 0 And VACIO% + APRO% = 10 Then Text7 = "A"

''   If UCase(TRIM$(Text5(0))) = "A" And TRIM$(Text1(1)) = "" And TRIM$(Text3(1)) = "" And TRIM$(Text5(1)) = "" Then
''       Text7 = "A"
''     ElseIf UCase(TRIM$(Text5(1))) = "A" And TRIM$(Text1(0)) = "" And TRIM$(Text3(0)) = "" And TRIM$(Text5(0)) = "" Then
''       Text7 = "A"
''   End If
   '
End Sub


Private Sub Text5_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus

End Sub

Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus

End Sub

Private Sub Text8_Change()
   '\\\OT-08-0935-OD-08/10/08///
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Visible = True
   Image1.Enabled = False
   Command23.Enabled = True
   '
TXTRAZA_COPUTIE07$ = ""
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 1 = " & Now & vbCrLf
   Call LIMPIAINFORMACION
TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 4 = " & Now & vbCrLf
  
  '\\\OT-08-0940-OD-FIN///
End Sub

Private Sub Text8_DblClick()
    '
    Call SELECHO("INDINUSE")
    TTXX$ = VALACT1$("INDINUSE")
    If TTXX$ <> "" Then
      Text8 = TTXX$
      Call Text8_KeyPress(13)
    End If
    
    '
End Sub
'
Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label13) ' SI VIENE DE IMPRESION SIMPLE TOMA DEL LABEL
    If N% > 0 Then
        REVIS% = 0 ' SI VIENE DE GRABACION BUSCA EL  VALOR MAXIMO
        NUSERIE$ = TRIM$(Text8)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR070 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        Dim MAXREVTMP As ADODB.Recordset
        'Set MAXREVTMP = readset(SQLZ$)
        Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    '
    TIPODOC$ = "Puesta a Tierra" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Text8) & "-" & REVIS%

    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Text11) 'FECHA
    CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = Label88.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(3) = "REF-MAT1":  DOCPARAM$(3) = TRIM$(Text8.TEXT) 'NRO SERIE
    CODPARAM$(4) = "REF-MAT2":  DOCPARAM$(4) = UCase(TRIM$(Text7.TEXT))  'VEREDICTO
    CODPARAM$(5) = "REF-MAT3":  DOCPARAM$(5) = TRIM$(Label11)  'Inspector
    CODPARAM$(6) = "OBSERVA":  DOCPARAM$(6) = TRIM$(Text9.TEXT)  'Observaciones
    CAPARDOC% = 6
    '
    CODTABU1$(1) = "N-PLANO"
    CODTABU1$(2) = "R-PLANO"
    CODTABU1$(3) = "F-PLANO"
    CODTABU1$(4) = "Z-PLANO"

    CACOLTAB1% = 4
    '
    CAITAB1% = 0
    For U& = 1 To 2
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Text1(U& - 1))    'CORRIENTE DE ENSAYO APLICADA
        TETABU1$(2, CAITAB1%) = TRIM$(Text3(U& - 1))   ' CAIDA DE TENSION MEDIDA
        TETABU1$(3, CAITAB1%) = TRIM$(Label9(U& - 1))    'RESISTENCIA CALCULADA
        TETABU1$(4, CAITAB1%) = UCase(TRIM$(Text5(U& - 1)))   ' VEREDICTO
    Next U&
    
    '
    FORIPRE$ = TRIM$(CONTROL$("", "INFENSA"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    For i& = 1 To 10
        Text1(i& - 1) = ""
        Text3(i& - 1) = ""
        Text5(i& - 1) = ""
        Label9(i& - 1) = ""
    Next i&
    Text2 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNICO
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Label88 = ""
    Label95 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA- BORRO LABEL DE REVISON
    Text11 = FECHATEX$(HOY(HO$))
    Label98 = ""
End Sub
Sub LIMPIAINFORMACION()
    For i& = 1 To 10
        Text1(i& - 1) = ""
        Text3(i& - 1) = ""
        Text5(i& - 1) = ""
        Label9(i& - 1) = ""
    Next i&
    Text2 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNICO
    Text7 = ""
    Text9 = ""
    Label88 = ""
    Label95 = ""
    Label98 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA- BORRO LABEL DE REVISON

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   
       NUSERIE$ = TRIM$(Text8)
       If InStr(NUSERIE$, "'") < 1 Then
            'VALIDACION DE NUMERO DE SERIE
            '
            CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
            If CANTRE& < 1 Then
               Call MENSERR(24, "Número de Serie Incorrecto")
               Exit Sub
            End If

           CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NumeroSerie ='" & NUSERIE$ & "'"))
           If CI% > 0 Then
            Label95 = CODEXT$(CI%)
            Label88 = DESCRIT0$(CI%)
           End If
           '\\\OT-08-0935-OD-FIN///
           '\\\OT-08-0940-OD-20/10/08///
'           REVIS% = 0
'           SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR070 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'           SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'           Dim ENSA070TMP As ADODB.Recordset
'           'Set ENSA070TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'           Set ENSA070TMP = READSET(SQLZ$)
'           ULTREVIS% = XVALO(ENSA070TMP!MAXREVI)
'           ENSA070TMP.Close
'           Set ENSA070TMP = Nothing
'           Label13 = ULTREVIS%
    TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 2 = " & Now & vbCrLf
           '
           BLQUEO$ = ""
           SQLX$ = "SELECT TOP 1 * FROM ENSAR070 WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
           'SQLX$ = SQLX$ + " AND REVISION = " & ULTREVIS% & "" '\\\OT-08-1064-OD-21/11/08-MEDITEA- INCLUIR LA  BUSQUEDA POR EL NUMERO DE REVISION
           SQLX$ = SQLX$ + " ORDER BY REVISION DESC "
    '       Dim ENSATMP As ADODB.Recordset
    '       Set ENSATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
           Set ENSATMP = READSET(SQLX$)
    TXTRAZA_COPUTIE07$ = TXTRAZA_COPUTIE07$ & "PUNTO 3 = " & Now & vbCrLf
           
           With ENSATMP
                If Not .EOF Then
                    Label13 = XVALO(!Revision)
                    CI% = !MODELO
                    Label95 = CODEXT$(CI%)
                    Label88 = DESCRIT0$(CI%)
                    Label11 = SAFETEXT$(!Inspector)
                    Text9 = SAFETEXT$(!Comentarios)
                    Label98 = SAFETEXT$(!INSTRUMENTO)
                    Text1(0) = FORMATNUM(XVALO(!Corr_Ensa1), "F6.2")
                    Text3(0) = FORMATNUM(XVALO(!CTension1), "F6.2")
                    Text5(0) = TRIM$(SAFETEXT$(!Veredicto1))
                    '
                    Text1(1) = FORMATNUM(XVALO(!Corr_Ensa2), "F6.2")
                    Text3(1) = FORMATNUM(XVALO(!CTension2), "F6.2")
                    Text5(1) = SAFETEXT$(!Veredicto2)
                    '
                    Text1(2) = FORMATNUM(XVALO(!Corr_Ensa3), "F6.2")
                    Text3(2) = FORMATNUM(XVALO(!CTension3), "F6.2")
                    Text5(2) = SAFETEXT$(!Veredicto3)
                    '
                    Text1(3) = FORMATNUM(XVALO(!Corr_Ensa4), "F6.2")
                    Text3(3) = FORMATNUM(XVALO(!CTension4), "F6.2")
                    Text5(3) = SAFETEXT$(!Veredicto4)
                    '
                    Text1(4) = FORMATNUM(XVALO(!Corr_Ensa5), "F6.2")
                    Text3(4) = FORMATNUM(XVALO(!CTension5), "F6.2")
                    Text5(4) = SAFETEXT$(!Veredicto5)
                    '
                    Text1(5) = FORMATNUM(XVALO(!Corr_Ensa6), "F6.2")
                    Text3(5) = FORMATNUM(XVALO(!CTension6), "F6.2")
                    Text5(5) = SAFETEXT$(!Veredicto6)
                    '
                    Text1(6) = FORMATNUM(XVALO(!Corr_Ensa7), "F6.2")
                    Text3(6) = FORMATNUM(XVALO(!CTension7), "F6.2")
                    Text5(6) = SAFETEXT$(!Veredicto7)
                    '
                    Text1(7) = FORMATNUM(XVALO(!Corr_Ensa8), "F6.2")
                    Text3(7) = FORMATNUM(XVALO(!CTension8), "F6.2")
                    Text5(7) = SAFETEXT$(!Veredicto8)
                    '
                    Text1(8) = FORMATNUM(XVALO(!Corr_Ensa9), "F6.2")
                    Text3(8) = FORMATNUM(XVALO(!CTension9), "F6.2")
                    Text5(8) = SAFETEXT$(!Veredicto9)
                    '
                    Text1(9) = FORMATNUM(XVALO(!Corr_Ensa10), "F6.2")
                    Text3(9) = FORMATNUM(XVALO(!CTension10), "F6.2")
                    Text5(9) = SAFETEXT$(!Veredicto10)

                    Label98 = SAFETEXT$(!INSTRUMENTO)
                    '
                    BLQUEO$ = SAFETEXT$(!Bloqueado)
                    Text7 = SAFETEXT$(!VerGeneral)
                    Text2 = SAFETEXT$(!TECNICO) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-AGREGAR LECTURA DEL TECNICO
                    fecccha$ = SAFETEXT$(!FECH_ENSA) ' si no tiene fecha pone la de hoy
                    If fecccha$ <> "" Then fff% = CVINT(fecccha$): Text11 = FECHATEX(fff%)
                    'si tiene fecha le damos formato "dd/mm/aa"
                    If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Text11.TEXT = HO$
               End If
          End With
          ENSATMP.Close
          Set ENSATMP = Nothing
      End If
      If BLQUEO$ = "SI" Then
            Image3.Visible = True
            Image1.Visible = False
            Image1.Enabled = True
            Command2.Enabled = False
         Else
            Command23.Enabled = False '\\\OT-08-1064-OD-21/11/08-MEIDTEA BLOQUEA LA IMPRESION
            Text2 = Label11
      End If

  End If
End Sub

Private Sub Text9_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus

End Sub
