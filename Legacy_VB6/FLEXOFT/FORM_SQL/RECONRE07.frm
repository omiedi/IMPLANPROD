VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RECONRE07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Reparaciones"
   ClientHeight    =   8340
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9075
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   8160
      ScaleHeight     =   8310
      ScaleWidth      =   990
      TabIndex        =   14
      Top             =   120
      Width           =   1020
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "RECONRE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Documentos Externos"
         Top             =   3360
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   116
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   117
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
         Height          =   650
         Left            =   120
         Picture         =   "RECONRE07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   120
         Picture         =   "RECONRE07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   840
         Visible         =   0   'False
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "RECONRE07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Graba e Imprime Registro"
         Top             =   1920
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   120
         Picture         =   "RECONRE07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Imprime Registro de Control Final"
         Top             =   6600
         Width           =   650
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "RECONRE07.frx":0F12
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2640
         Width           =   650
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   0
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   280
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   195
         TabIndex        =   87
         Top             =   840
         Visible         =   0   'False
         Width           =   280
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   3885
      TabIndex        =   28
      Top             =   840
      Width           =   4215
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   2760
         TabIndex        =   13
         Text            =   " "
         Top             =   4920
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   2760
         TabIndex        =   12
         Text            =   " "
         Top             =   4560
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   2760
         TabIndex        =   11
         Text            =   " "
         Top             =   4200
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   2760
         TabIndex        =   10
         Text            =   " "
         Top             =   3840
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   2760
         TabIndex        =   9
         Text            =   " "
         Top             =   3480
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   2760
         TabIndex        =   8
         Text            =   " "
         Top             =   3120
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2760
         TabIndex        =   7
         Text            =   " "
         Top             =   2760
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2760
         TabIndex        =   6
         Text            =   " "
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2760
         TabIndex        =   5
         Text            =   " "
         Top             =   2040
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Text            =   " "
         Top             =   1680
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Text            =   " "
         Top             =   1320
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "16"
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
         TabIndex        =   115
         Top             =   5775
         Width           =   300
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "17"
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
         TabIndex        =   114
         Top             =   6075
         Width           =   300
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "18"
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
         TabIndex        =   113
         Top             =   6435
         Width           =   300
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "19"
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
         TabIndex        =   112
         Top             =   6750
         Width           =   300
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "20"
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
         TabIndex        =   111
         Top             =   7050
         Width           =   300
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   600
         TabIndex        =   110
         Top             =   5775
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   109
         Top             =   5775
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   600
         TabIndex        =   108
         Top             =   6090
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   107
         Top             =   6090
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   600
         TabIndex        =   106
         Top             =   6420
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   105
         Top             =   6420
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   18
         Left            =   600
         TabIndex        =   104
         Top             =   6735
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   18
         Left            =   1080
         TabIndex        =   103
         Top             =   6735
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   19
         Left            =   600
         TabIndex        =   102
         Top             =   7050
         Width           =   360
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   19
         Left            =   1080
         TabIndex        =   101
         Top             =   7050
         Width           =   360
      End
      Begin VB.Label Label83 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   90
         Top             =   6945
         Width           =   2205
      End
      Begin VB.Label Label82 
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
         Height          =   300
         Left            =   1920
         TabIndex        =   89
         Top             =   6585
         Width           =   780
      End
      Begin VB.Line Line4 
         X1              =   1785
         X2              =   7735
         Y1              =   6405
         Y2              =   6405
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH12:"
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
         Left            =   2160
         TabIndex        =   86
         Top             =   4995
         Width           =   540
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH11:"
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
         Left            =   2160
         TabIndex        =   85
         Top             =   4635
         Width           =   540
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH10:"
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
         Left            =   2160
         TabIndex        =   84
         Top             =   4275
         Width           =   540
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH09:"
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
         Left            =   2160
         TabIndex        =   83
         Top             =   3915
         Width           =   540
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH08:"
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
         Left            =   2160
         TabIndex        =   82
         Top             =   3555
         Width           =   540
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH07:"
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
         Left            =   2160
         TabIndex        =   81
         Top             =   3195
         Width           =   540
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH06:"
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
         Left            =   2160
         TabIndex        =   80
         Top             =   2835
         Width           =   540
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH05:"
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
         Left            =   2160
         TabIndex        =   79
         Top             =   2475
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH04:"
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
         Left            =   2160
         TabIndex        =   78
         Top             =   2115
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH03:"
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
         Left            =   2160
         TabIndex        =   77
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH02:"
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
         Left            =   2160
         TabIndex        =   76
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CH01:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   75
         Top             =   1035
         Width           =   540
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   1800
         X2              =   1785
         Y1              =   480
         Y2              =   7455
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Nº Control         Mediciones   "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   105
         TabIndex        =   29
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   74
         Top             =   5440
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   14
         Left            =   600
         TabIndex        =   73
         Top             =   5440
         Width           =   360
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15"
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
         TabIndex        =   72
         Top             =   5470
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   71
         Top             =   5120
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   600
         TabIndex        =   70
         Top             =   5120
         Width           =   360
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "14"
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
         TabIndex        =   69
         Top             =   5170
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   1080
         TabIndex        =   68
         Top             =   4800
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   600
         TabIndex        =   67
         Top             =   4800
         Width           =   360
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "13"
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
         TabIndex        =   66
         Top             =   4860
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   65
         Top             =   4480
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   600
         TabIndex        =   64
         Top             =   4480
         Width           =   360
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12"
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
         TabIndex        =   63
         Top             =   4500
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   62
         Top             =   4160
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   600
         TabIndex        =   61
         Top             =   4160
         Width           =   360
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "11"
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
         TabIndex        =   60
         Top             =   4200
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   59
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   600
         TabIndex        =   58
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10"
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
         TabIndex        =   57
         Top             =   3900
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   56
         Top             =   3520
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   55
         Top             =   3520
         Width           =   360
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9"
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
         TabIndex        =   54
         Top             =   3550
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   53
         Top             =   3200
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   600
         TabIndex        =   52
         Top             =   3200
         Width           =   360
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "8"
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
         TabIndex        =   51
         Top             =   3250
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   50
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   600
         TabIndex        =   49
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7"
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
         TabIndex        =   48
         Top             =   2940
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   47
         Top             =   2560
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   600
         TabIndex        =   46
         Top             =   2560
         Width           =   360
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         TabIndex        =   45
         Top             =   2580
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   44
         Top             =   2240
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   600
         TabIndex        =   43
         Top             =   2240
         Width           =   360
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         TabIndex        =   42
         Top             =   2270
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   41
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   600
         TabIndex        =   40
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         TabIndex        =   39
         Top             =   1950
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   38
         Top             =   1600
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   37
         Top             =   1600
         Width           =   360
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         TabIndex        =   36
         Top             =   1650
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   35
         Top             =   1280
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   34
         Top             =   1280
         Width           =   360
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         TabIndex        =   33
         Top             =   1300
         Width           =   300
      End
      Begin VB.Label Lis2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   32
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Lis1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   31
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         TabIndex        =   30
         Top             =   990
         Width           =   300
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "ENCABEZADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   3615
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         TabIndex        =   93
         Text            =   " "
         Top             =   2880
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2160
         TabIndex        =   0
         Text            =   " "
         Top             =   2880
         Width           =   1305
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3600
         Width           =   3330
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1150
         TabIndex        =   100
         Top             =   1140
         Width           =   1700
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   99
         Top             =   600
         Width           =   1155
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
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label86 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2850
         TabIndex        =   92
         Top             =   1200
         Width           =   200
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   91
         ToolTipText     =   "Número de Revisión"
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Dictámen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   27
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tensión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   24
         Top             =   3340
         Width           =   1410
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   3345
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   1680
         Width           =   2640
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   21
         Top             =   1750
         Width           =   780
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   80
         TabIndex        =   19
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/F."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   1440
      End
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6600
      TabIndex        =   119
      Top             =   480
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5880
      Picture         =   "RECONRE07.frx":121C
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "RECONRE07.frx":1526
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00CDDADA&
      Caption         =   "Registro de Reparaciones"
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
      TabIndex        =   15
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "RECONRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim CantidadControlesArticulo As Integer

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    'VALIDACION POR FECHA
    
    VDEVV$ = TRIM$(Label12)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    'VALIDACION DE NUMERO DE SERIE
    '
    NUSERIE$ = TRIM$(Label24)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    'VALIDACION DE DICTAMEN
    If TRIM$(Text5) = "" Then
      Call MENSERR(24, "Falta Informar Dictamen")
      GoTo 99
    End If
    '
    If TRIM$(Label8) = "" Then
      Call MENSERR(24, "Falta Tensión del Equipo")
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "NumeroSerie = '" & NUSERIE$ & "'"
    '
    '
    Call ACTUREGISTRO("TANUMSE", "Dic_Ensa", SQLX$, Left$(TRIM$(Text5), 4), REGAF&)
    Call ACTUREGISTRO("TANUMSE", "Obse_Ensa", SQLX$, Text4, REGAF&)
    Call ACTUREGISTRO("TANUMSE", "Fech_Ensa", SQLX$, TRIM$(Label12), REGAF&)
    '
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Registro de Control Final de Reparaciones" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Label83))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text5))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    REVIS% = 0
    '
    CONDI$ = " NumeroSerie ='" & NUSERIE$ & "'  and ORIGEN = 'REP' "
    ULTREVIS% = XVALO(VALOBADA("CONTROLFIN", "MAX(REVISION)", CONDI$, "NOMESS"))
    Label85 = ULTREVIS%
    '
    SQLX$ = "SELECT * FROM CONTROLFIN "
    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
    SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
    '
    Dim CONTRTMP As ADODB.Recordset
    Set CONTRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CONTRTMP
     If Not .EOF Then
       REVIS% = 1 + ULTREVIS%
     Else
       REVIS% = 0 'SI ES EL PRIMERO EL VALOR VA A SER 0
     End If
    End With
    CONTRTMP.Close
    Set CONTRTMP = Nothing
    '
    'GRABO LOS DATOS COMO NUEVOS EN CONTROLFIN
    SQLX$ = "SELECT * FROM CONTROLFIN"
    SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NUSERIE$ & "'"
    Dim CTRLFINTMP As ADODB.Recordset
    Set CTRLFINTMP = New ADODB.Recordset
    CTRLFINTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For U& = 0 To Lis1.UBound
        With CTRLFINTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUMEROSERIE = NUSERIE$
          !COD_INTE = CODINT%(Label7)
          !NumeroOrden = U& + 1
          RESTDO$ = ""
          If Lis1(U&) <> "" Then RESTDO$ = "OK"
          If Lis2(U&) <> "" Then RESTDO$ = "NC"
          !Resultado = RESTDO$
          '!MARBORRA = 0
          !Fech_Ensa = TRIM$(Label12)
          !Revision = REVIS%
          !ORIGEN = "REP"
          .Update
        End With
    Next U&
    CTRLFINTMP.Close
    '
    '
    'GRABO LOS DATOS COMO NUEVOS EN MEDIFIN
    SQLX$ = "SELECT * FROM MEDIFIN"
    SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NUSERIE$ & "'"
    Dim MEDIFINTMP As ADODB.Recordset
    Set MEDIFINTMP = New ADODB.Recordset
    MEDIFINTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For i& = 0 To 11
        With MEDIFINTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUMEROSERIE = NUSERIE$
          !COD_INTE = CODINT%(Label7)
          TTT$ = TRIM$(Str(i& + 1))
          While Len(TTT$) < 2: TTT$ = "0" + TTT$: Wend
          !CHN = "CH" + TTT$
          !MEDA = XVALO(TextA(i&))
          !Fech_Ensa = TRIM$(Label12)
          !Revision = REVIS%
          !ORIGEN = "REP"
          .Update
        End With
    Next i&
    MEDIFINTMP.Close
    
    '
    Call PRINTMEDI(1)
    '
    Command23.Enabled = True
    Call BLANFORMU
    '
    Call Command1_Click
    
99  Screen.MousePointer = 1
    '
    Command2.Enabled = True
    '
    
End Sub

Private Sub Command23_Click()
    NUSERIE$ = TRIM$(Label24)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       Exit Sub
    End If
    CANCELPRINT% = 1
    Call PRINTMEDI(0) '  0 IMPRESION SIMPLE --- UNO GRABACION
    CANCELPRINT% = 0
    
    FILTX$ = TRIM$(Label45) & ";" & Label85
    Call STDFORM("ORCF-MDT", FILTX$, "PRINT")

End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPRCF.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command4_Click()
   
   If XVALO(VALOBADA("TANUMSE", "NumeroSerie", "NumeroSerie ='" & TRIM$(Label24) & "'")) > 0 Then
     Call VERIMPENSAYOS
   End If
End Sub



Private Sub Command7_Click()

    Opcion% = COMALTER%("Documentos Externos\\Visualizar\Agregar Nuevos")
    
    If Opcion% = 1 Then
        TituloLista$ = "DOCUMENTOS EXTERNOS"
        
        CAMPOS$ = "Nombre/A48;Ruta/A0"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
        
        IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
        SQLX$ = "SELECT CAMPO4, CAMPO5 FROM DATASQL WITH(NOLOCK) WHERE ID_LISTA=" & CStr(IDLIST&) & " AND CAMPO2='" & Label7 & "'"
        Set RS = READSET(SQLX$)
        With RS
            If Not (.EOF And .BOF) Then
                FIN& = RS.RecordCount
                FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
            End If
            Do While Not .EOF
                NombreDocumento$ = !CAMPO4
                RutaDocumento$ = !CAMPO5
                
                J& = J& + 1
                FRMZELECHO.msf2.TextMatrix(J&, 1) = NombreDocumento$
                FRMZELECHO.msf2.TextMatrix(J&, 2) = RutaDocumento$
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RS.Close
        Set RS = Nothing
        '
        On Error GoTo 0
        
        Screen.MousePointer = 1
        
        FRMZELECHO.Width = 8000
50      FRMZELECHO.Show 1
    
        RutaDocumento$ = TRIM$(RESP_ZELE_VECT(2))
        If RutaDocumento$ <> "" Then
            Call ShellExecute(Me.hwnd, "Open", RutaDocumento$, "", "", 1)
        Else
            GoTo 99
        End If
            
        GoTo 50
    
    ElseIf Opcion% = 2 Then
        
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Todos los Archivos(*.*)|*.*|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        RutaDocumento$ = TRIM$(UCase$(CommonDialog1.FileName))
        
        If RutaDocumento$ <> "" Then
            NombreDocumento$ = TRIM$(InputBox("Nombre", "Ingrese Nombre para el Documento"))
            
            If NombreDocumento$ <> "" Then
                CodigoExterno$ = Label7
                DescripcionProducto$ = Label13
                
                IDLIST& = ObtenerIdTablaAuxiliar("DOCEXCONFIN")
                
                ConsultaSql$ = "INSERT INTO DATASQL (ID_LISTA, CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5) VALUES ('" & CStr(IDLIST&) & "','','" & CodigoExterno$ & "','" & DescripcionProducto$ & "','" & NombreDocumento$ & "','" & RutaDocumento$ & "')"
                FLEXCONN.Execute ConsultaSql$
            End If
        End If

    End If
    
99  Exit Sub


End Sub

Private Sub Form_Load()
    '
'    Call BLANFORMU
    '
    FF% = HOY(HO$)
    Label12 = HO$
    Label83 = USERNAME$
    '
End Sub


Private Sub Image1_DblClick()
   If DERACCE("MODENS70", "Modificacion de Informe de Ensayo") > 1 Then
        Image3.Visible = True
        Image1.Visible = False
        Command2.Enabled = False
        Command23.Enabled = True ' HABILITA LA IMPRESORA
        ULTREVIS% = XVALO(Label85)
        'SI YA TIENE ALGUNA REVISION MUESTRA LA FECHA DE LA REVISION
        If ULTREVIS% > 0 Then
          FEC = VALOBADA("TANUMSE", "FECH_ENSA", "NumeroSerie ='" & Label24 & "'")
          FE% = CVINT(FEC) ' SI VUELVE A ESTAR EL  CANDADO VISIBLE PONGO
          Label12 = FECHATEX$(FE%) ' LA FECHA QUE TIENE EL ENSAYO GRABADA
        Else 'SI ES LA PRIMERA REVISION MUESTRA LA DEL DIA
          HOII% = HOY(HOS$)
          Label12 = FECHATEX$(HOII%)
        End If
   End If
End Sub

Private Sub Image3_DblClick()
   If DERACCE("MODCONFI", "Modificacion  de Registro de Control Final") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False
        HOII% = HOY(HOS$)
        Label12 = HOS$ ' PONGO LA FECHA ACTUAL
   End If
End Sub

Private Sub Label8_DblClick()
   If Image3.Visible = True Then Exit Sub 'candado visible sale de la rutina
   If DERACCE("CAMBITEN", "Cambio de Tension") > 1 Then
     Text8.Visible = True 'muestra el text de tension y tapa al label
   End If
End Sub

Private Sub Lis1_DblClick(Index As Integer)
  If Image3.Visible = True Then Text5.SetFocus: Exit Sub

   '\\\OT-08-0933-OD-06/10/08///
    If Lis1(Index) = "" Then
       Lis1(Index) = "OK"
       Lis2(Index) = ""
     Else
       Lis1(Index) = ""
    End If
   '\\\OT-08-0933-OD-FIN///
   Call VeredGeneral 'MEDITEA-VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN
End Sub

Private Sub Lis2_DblClick(Index As Integer)
   '\\\OT-08-0933-OD-06/10/08///
    If Image3.Visible = True Then Text5.SetFocus: Exit Sub

    If Lis2(Index) = "" Then
       Lis2(Index) = "NC"
       Lis1(Index) = ""
     Else
       Lis2(Index) = ""
    End If
   
    Call VeredGeneral 'VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN

End Sub

Private Sub Text1_GotFocus()
  If Image3.Visible = True Then Text5.SetFocus
End Sub



Private Sub Label24_Change()
   '\\\OT-08-0933-OD-06/10/08///+
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Enabled = False
   Image1.Visible = True
   Command23.Enabled = True
   NUSERIE$ = TRIM$(Label24)
   
   'Call LIMPIAINFORMACION
   If InStr(NUSERIE$, "'") < 1 And NUSERIE$ <> "" Then

       SQLX$ = "SELECT * FROM TANUMSE"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       Dim TNMSETMP As ADODB.Recordset
       Set TNMSETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With TNMSETMP
            If Not .EOF Then
                
                CI% = XVALO(!COD_INTE)
                Label15 = SAFETEXT$(!IDSUBOR)
                Label7 = CODEXT$(CI%)
                Label13 = DESCRIT0$(CI%)
                Text4 = SAFETEXT$(!Obse_Ensa)
                Text5 = SAFETEXT$(!Dic_Ensa)
                Label8 = SAFETEXT$(!Ten_Ensa)
                Text8 = XVALO(!Ten_Ensa)
                fecccha$ = SAFETEXT$(!Fech_Ensa) ' si no tiene fecha pone la de hoy
                If fecccha$ <> "" Then fff% = CVINT(fecccha$): Label12 = FECHATEX(fff%)
                If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Label12 = HO$
    
                GoSub CargaRestDatos
            End If
      End With
CargaRestDatos:
       '
       CONDI$ = " NumeroSerie ='" & NUSERIE$ & "'  and ORIGEN = 'REP' "
       ULTREVIS% = XVALO(VALOBADA("CONTROLFIN", "MAX(REVISION)", CONDI$, "NOMESS"))
       Label85 = ULTREVIS%
       Label83 = VALOBADA("REPARA", "UsRealizaRepa", "NUMREPA = " & Label45)
       'AGREGAR VALIDACION SI ES LA PRIMERA REVISION QUE PRESENTE FECHA DE HOY
       If ULTREVIS% < 1 Then
          HOII% = HOY(HOS$)
          Label12 = FECHATEX$(HOII%)
       End If
       '
'       If ULTREVIS% > 0 Then ' SI TIENE REVISION DE REPARACION MUESTRA SI NO NO
       SQLX$ = "SELECT * FROM CONTROLFIN WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
       SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
       Dim CTRLFINTMP As ADODB.Recordset
       Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With CTRLFINTMP
        If Not .EOF Then
         Call bloquear
         Do While Not .EOF
           IND% = XVALO(!NumeroOrden)
           If !Resultado = "OK" Then Lis1(IND% - 1) = SAFETEXT$(!Resultado)
           If !Resultado = "NC" Then Lis2(IND% - 1) = SAFETEXT$(!Resultado)
           If !Resultado = "" Then Lis2(IND% - 1) = ""
         .MoveNext
         Loop
        Else
         Command23.Enabled = False
        End If
       End With
       '
       J& = 0
       SQLX$ = "SELECT * FROM MEDIFIN WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
       SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
       SQLX$ = SQLX$ + " AND ORIGEN = 'REP' "
       SQLX$ = SQLX$ + " ORDER BY CHN "

       Dim MEDIFINTMP As ADODB.Recordset
       Set MEDIFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MEDIFINTMP
        If Not .EOF Then
         Call bloquear
       
         Do While Not .EOF
           If J& > 11 Then Exit Do
           TextA(J&) = FORMATNUM$(XVALO(!MEDA), "F7.2")  '
           J& = J& + 1
         .MoveNext
         Loop
        End If
       End With
       
       IDLIST& = ObtenerIdTablaAuxiliar("CANCONART")
       Codigo$ = TRIM$(Label7)
       CantidadControlesArticulo = XVALO(VALOBADA("DATASQL", "CAMPO3", "ID_LISTA=" & CStr(IDLIST&) & " AND CAMPO1='" & Codigo$ & "'"))
       
   End If
'   End If
   '\\\OT-08-0933-OD-FIN///
End Sub


Private Sub Text4_GotFocus()
   If Image3.Visible = True Then Text5.SetFocus
End Sub


Private Sub Text6_GotFocus()
   If Image3.Visible = True Then Text5.SetFocus
End Sub


Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Text5.SetFocus
End Sub

Private Sub Text8_GotFocus()
   If Image3.Visible = True Then Text5.SetFocus '
End Sub

Private Sub TextA_GotFocus(Index As Integer)
   If Image3.Visible = True Then Text5.SetFocus
End Sub

Private Sub TextA_LostFocus(Index As Integer)
   
  VALOR1 = ROUND#(XVALO(TextA(Index)), 2)

  If XVALO(VALOR1) < 10000 Then
   TextA(Index) = FORMATNUM(TextA(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    TextA(Index).SetFocus
  End If
   
End Sub
Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label85)
    If N% > 0 Then
        REVIS% = 0
        NUSERIE$ = TRIM$(Label24)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM CONTROLFIN"
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        Dim MAXREVTMP As ADODB.Recordset
        Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    TIPODOC$ = "Registro de Reparacion" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Label24) & "-" & REVIS%
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Label12) 'FECHA
    CODPARAM$(2) = "NRO-AENT": DOCPARAM$(2) = Label15 'o.f
    CODPARAM$(3) = "TALLE": DOCPARAM$(3) = TRIM$(Label83) 'TECNICO
    'CODPARAM$(4) = "COLOR": DOCPARAM$(4) = TRIM$(Text7) 'A.EXTERNO
    CODPARAM$(5) = "REF-MAT1":  DOCPARAM$(5) = TRIM$(Label24) 'NRO SERIE
    CODPARAM$(6) = "DES-MAT": DOCPARAM$(6) = Label13.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(7) = "REF-MAT2":  DOCPARAM$(7) = Text5.TEXT  'DICTAMEN
    CODPARAM$(8) = "OBSERVA":  DOCPARAM$(8) = Text4.TEXT  'OBSERVACIONES
    CODPARAM$(9) = "N-HPROF":  DOCPARAM$(9) = XVALO((Text8.TEXT)) '\\\OT-08-1074-OD-24/11/08-MEDITEA-TENSION
    CAPARDOC% = 10
    '
    CODTABU1$(1) = "MEDIDA-1"
    CODTABU1$(2) = "MEDIDA-2"
    CODTABU1$(3) = "MEDIDA-3"
    CODTABU1$(4) = "N-PLANO"
    CODTABU1$(5) = "R-PLANO"
    CODTABU1$(6) = "F-PLANO"
    CODTABU1$(7) = "Z-PLANO"
    CACOLTAB1% = 7
    '
    U& = 0
    J& = 0
    CAITAB1% = 0
    For U& = 0 To 11
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(TextA(U&))     'MEDICIONES A
    Next U&
    CAITAB1% = 0
    For J& = 0 To Lis1.UBound
        CAITAB1% = CAITAB1% + 1
        TETABU1$(4, CAITAB1%) = TRIM$(Lis1(J&))   ' NUMERO
        TETABU1$(5, CAITAB1%) = TRIM$(Lis2(J&))    'CONTROL
    Next J&
    '
    FORIPRE$ = TRIM$(Control("", "RECONRE"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    i& = 0
    For i& = 0 To 14
        Lis1(i&) = ""
        Lis2(i&) = ""
    Next i&
    i& = 0
    For i& = 0 To 11
        TextA(i&) = ""
    Next i&
    Label15 = ""
    Label7 = ""
    Label13 = ""
    Text5 = ""
    Label8 = ""
    Text4 = ""
    Label12 = FECHATEX$(HOY(HO$))
    Label85 = ""
    Text8.Visible = False
    Label12 = ""
    Label24 = ""
End Sub




Sub LIMPIAINFORMACION()
    Label8 = ""
    Label13 = ""
    Label15 = ""
    Label7 = ""
    Text4 = ""
    Text5 = ""
    i& = 0
    For i& = 0 To 14
        Lis1(i&) = ""
        Lis2(i&) = ""
    Next i&
    i& = 0
    For i& = 0 To 11
        TextA(i&) = ""
    Next i&
    Label85 = ""
    Text8.Visible = False
End Sub
Sub bloquear()
    Image3.Visible = True
    Image1.Visible = False
    Command2.Enabled = False
    Image1.Enabled = True

End Sub
Sub VeredGeneral()
  OK% = 0
  NC% = 0
  For i& = 0 To Lis2.UBound
    If Lis2(i&) = "NC" Then
      Text5 = "R"
      Exit Sub
    End If
  Next i&

  For i& = 0 To Lis1.UBound
    If Lis1(i&) = "OK" Then
      OK% = OK% + 1
    End If
  Next i&
  
  If OK% = Lis1.UBound + 1 Or (OK% >= CantidadControlesArticulo And CantidadControlesArticulo > 0) Then
    Text5 = "A"
  Else
    Text5 = ""
  End If
End Sub
Sub VERIMPENSAYOS()
   Command4.Enabled = False
   '
   '
   Call COPYSTRU("INDIENSA", "INDIENRE")
   Call SETULTREG("!INDIENRE", 0)

   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM CONTROLFIN "
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Label24 & "'"
   SQLX$ = SQLX$ + " AND ORIGEN = 'REP'"
   Dim CTRLFINTMP As ADODB.Recordset
   Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With CTRLFINTMP
     U& = 0
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!Fech_Ensa)
       NUREV$ = "Registro de Reparacion " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENRE", TX$, U&)
       .MoveNext
     Loop
   End With
   CTRLFINTMP.Close
   Set CTRLFINTMP = Nothing
   '
   '
   Call SETULTREG("!INDIENRE", U&)
   Call FRESHECHO("!INDIENRE")
   '
10 If ULTREG&("!INDIENRE") > 0 Then
     Call SELECHO("!INDIENRE/Indice de Registro de Reparaciones")
   Else
     GoTo 99
   End If
   '
   NUDOCU$ = VALACT2$("!INDIENRE") ' MUESTRO EL FORMULARIO IMPRESO  SELECCIONAD0
   If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
    POS% = InStr(NUDOCU$, TRIM$(Label24)) ' ASIGNO A POS% LA POSICION DESDE DONDE EMPIEZA EL NUMERO DE SEREIE
    TIDOCUM$ = TRIM$(Mid$(NUDOCU$, 1, POS% - 1)) ' ASIGNO A TIDOCUM SOLO EL TIPO DE ENSAYO QUE VA A SER DESDE 1 HASTA EL VALOR DE POS% - 1
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
           PPXX% = InStr(XX$, "Nro.")
           If PPXX% > 0 Then
             XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
           End If
           If InStr(XX$, NUDOCU$) > 4 Then
             DOCUNU& = CVS(Left$(XX$, 4))
             GoTo 5
           End If
         End If
    Next UI&
       '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Ensayo.\ ")
    GoTo 10
       '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
   GoTo 10
99 Command4.Enabled = True
End Sub
 
