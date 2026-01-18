VERSION 5.00
Begin VB.Form FORSELEC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0DD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8805
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5565
      Top             =   1050
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   960
      Left            =   -105
      ScaleHeight     =   900
      ScaleWidth      =   8880
      TabIndex        =   157
      Top             =   3675
      Visible         =   0   'False
      Width           =   8940
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8085
         TabIndex        =   164
         Top             =   0
         Width           =   225
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00004040&
         Caption         =   "Tipo de Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   680
         Left            =   210
         TabIndex        =   162
         Top             =   50
         Width           =   2325
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   250
            Width           =   2010
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00004040&
         Caption         =   "Grupo Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   680
         Left            =   2625
         TabIndex        =   160
         Top             =   50
         Width           =   2325
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   250
            Width           =   2010
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00004040&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   680
         Left            =   5040
         TabIndex        =   158
         Top             =   50
         Width           =   2955
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   250
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton BOTFIL 
      Caption         =   "Filter"
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
      Left            =   3150
      TabIndex        =   156
      Top             =   2910
      Width           =   600
   End
   Begin VB.ListBox LISORT 
      Height          =   540
      ItemData        =   "FORSELEC.frx":0000
      Left            =   5460
      List            =   "FORSELEC.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   154
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   15
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   15
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   15
         Left            =   105
         TabIndex        =   144
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   15
         Left            =   3750
         TabIndex        =   143
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   15
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   15
         Left            =   4200
         TabIndex        =   150
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   15
         Left            =   4200
         TabIndex        =   149
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   15
         Left            =   4200
         TabIndex        =   148
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   15
         Left            =   4200
         TabIndex        =   147
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   15
         Left            =   4200
         TabIndex        =   146
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   14
      Left            =   0
      TabIndex        =   133
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   14
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   14
         Left            =   105
         TabIndex        =   135
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   14
         Left            =   3750
         TabIndex        =   134
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   14
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   14
         Left            =   4200
         TabIndex        =   141
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   14
         Left            =   4200
         TabIndex        =   140
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   14
         Left            =   4200
         TabIndex        =   139
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   14
         Left            =   4200
         TabIndex        =   138
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   14
         Left            =   4200
         TabIndex        =   137
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   13
      Left            =   0
      TabIndex        =   124
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   13
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   13
         Left            =   105
         TabIndex        =   126
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   13
         Left            =   3750
         TabIndex        =   125
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   13
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   13
         Left            =   4200
         TabIndex        =   132
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   13
         Left            =   4200
         TabIndex        =   131
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   13
         Left            =   4200
         TabIndex        =   130
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   13
         Left            =   4200
         TabIndex        =   129
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   13
         Left            =   4200
         TabIndex        =   128
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   12
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   12
         Left            =   0
         TabIndex        =   118
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   12
         Left            =   105
         TabIndex        =   117
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   12
         Left            =   3750
         TabIndex        =   116
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   12
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   12
         Left            =   4200
         TabIndex        =   123
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   12
         Left            =   4200
         TabIndex        =   122
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   12
         Left            =   4200
         TabIndex        =   121
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   12
         Left            =   4200
         TabIndex        =   120
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   12
         Left            =   4200
         TabIndex        =   119
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   11
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   11
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   11
         Left            =   105
         TabIndex        =   108
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   11
         Left            =   3750
         TabIndex        =   107
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   11
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   11
         Left            =   4200
         TabIndex        =   114
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   11
         Left            =   4200
         TabIndex        =   113
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   11
         Left            =   4200
         TabIndex        =   112
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   11
         Left            =   4200
         TabIndex        =   111
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   11
         Left            =   4200
         TabIndex        =   110
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   10
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   10
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   10
         Left            =   105
         TabIndex        =   99
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   10
         Left            =   3750
         TabIndex        =   98
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   10
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   105
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   104
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   103
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   102
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   101
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   9
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   9
         Left            =   0
         TabIndex        =   91
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   9
         Left            =   105
         TabIndex        =   90
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   9
         Left            =   3750
         TabIndex        =   89
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   9
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   9
         Left            =   4200
         TabIndex        =   96
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   9
         Left            =   4200
         TabIndex        =   95
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   9
         Left            =   4200
         TabIndex        =   94
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   9
         Left            =   4200
         TabIndex        =   93
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   9
         Left            =   4200
         TabIndex        =   92
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   8
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   8
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   8
         Left            =   105
         TabIndex        =   81
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   8
         Left            =   3750
         TabIndex        =   80
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   8
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   87
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   86
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   85
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   84
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   83
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   7
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   7
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   7
         Left            =   105
         TabIndex        =   72
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   7
         Left            =   3750
         TabIndex        =   71
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   7
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   7
         Left            =   4200
         TabIndex        =   78
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   7
         Left            =   4200
         TabIndex        =   77
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   7
         Left            =   4200
         TabIndex        =   76
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   7
         Left            =   4200
         TabIndex        =   75
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   7
         Left            =   4200
         TabIndex        =   74
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   6
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   6
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   6
         Left            =   105
         TabIndex        =   63
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   6
         Left            =   3750
         TabIndex        =   62
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   6
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   6
         Left            =   4200
         TabIndex        =   69
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   6
         Left            =   4200
         TabIndex        =   68
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   6
         Left            =   4200
         TabIndex        =   67
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   6
         Left            =   4200
         TabIndex        =   66
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   6
         Left            =   4200
         TabIndex        =   65
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   5
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   5
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   5
         Left            =   105
         TabIndex        =   54
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   5
         Left            =   3750
         TabIndex        =   53
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   5
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   5
         Left            =   4200
         TabIndex        =   60
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   5
         Left            =   4200
         TabIndex        =   59
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   5
         Left            =   4200
         TabIndex        =   58
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   5
         Left            =   4200
         TabIndex        =   57
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   5
         Left            =   4200
         TabIndex        =   56
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   4
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   4
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   4
         Left            =   105
         TabIndex        =   45
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   4
         Left            =   3750
         TabIndex        =   44
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   4
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   4
         Left            =   4200
         TabIndex        =   51
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   4
         Left            =   4200
         TabIndex        =   50
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   4
         Left            =   4200
         TabIndex        =   49
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   4
         Left            =   4200
         TabIndex        =   48
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   4
         Left            =   4200
         TabIndex        =   47
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   3
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   3
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   3
         Left            =   105
         TabIndex        =   36
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   3
         Left            =   3750
         TabIndex        =   35
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   3
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   42
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   41
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   40
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   39
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   38
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   2
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   2
         Left            =   105
         TabIndex        =   27
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   2
         Left            =   3750
         TabIndex        =   26
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   2
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   2
         Left            =   4200
         TabIndex        =   33
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   2
         Left            =   4200
         TabIndex        =   32
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   2
         Left            =   4200
         TabIndex        =   31
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   2
         Left            =   4200
         TabIndex        =   30
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   2
         Left            =   4200
         TabIndex        =   29
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5355
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   1
         Left            =   3750
         TabIndex        =   17
         Top             =   0
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   1
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   24
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   23
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   22
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   21
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   20
         Top             =   1575
         Width           =   855
      End
   End
   Begin VB.CommandButton BOTINSE 
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
      Height          =   360
      Left            =   3465
      TabIndex        =   3
      Top             =   2910
      Width           =   430
   End
   Begin VB.CommandButton BOTERA 
      Caption         =   "Del"
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
      Left            =   3885
      TabIndex        =   4
      Top             =   2910
      Width           =   430
   End
   Begin VB.Frame MARCOTA 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2760
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5355
      Begin VB.VScrollBar VScroll1 
         Height          =   2745
         Index           =   0
         Left            =   3750
         TabIndex        =   10
         Top             =   0
         Width           =   225
      End
      Begin VB.ListBox TASEL 
         BackColor       =   &H00D0FFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ListBox TABLA 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Label FILTER 
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   15
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label ATABLA 
         Caption         =   "ATABLA"
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   14
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label RENUAC 
         Caption         =   "RENUAC"
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   13
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD2 
         Caption         =   "VD2"
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   12
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label VD1 
         Caption         =   "VD1"
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Index           =   0
         X1              =   3710
         X2              =   3710
         Y1              =   0
         Y2              =   2730
      End
   End
   Begin VB.CommandButton BOTCANCE 
      Caption         =   "X"
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
      Left            =   4305
      TabIndex        =   5
      Top             =   2910
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      TabIndex        =   0
      Top             =   2910
      Width           =   2115
   End
   Begin VB.Label Label99 
      Caption         =   "Microsoft Vísual Basic 5.O ................................"
      Height          =   225
      Left            =   0
      TabIndex        =   155
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label ABUSQUE 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6090
      TabIndex        =   153
      Top             =   2940
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LINI 
      Caption         =   "LINI"
      Height          =   225
      Left            =   5460
      TabIndex        =   152
      Top             =   2310
      Width           =   645
   End
   Begin VB.Label SINI 
      Caption         =   "SINI"
      Height          =   225
      Left            =   5460
      TabIndex        =   151
      Top             =   1995
      Width           =   645
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3990
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label ARTITUL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5250
      TabIndex        =   6
      Top             =   2940
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
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
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Doble-Click para Filtrado Completo"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   1
      Top             =   1785
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "FORSELEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHITA$             ' archivo tabla activa
Dim TITULTA$             ' titulo tabla activa
Dim VECARTA$             ' string indice de tablas
Dim ITABLAC%             ' indice tabla activa 0 ... 15
Dim REFECHAC&
Dim OFSEVER&
'
Dim TV1%(15), TV2%(15)
Dim LF1%(16), LF2%(15)
Dim PR1%(15), PR2%(15)
Dim JT1%(15), JT2%(15)
Dim LD1%(15), LD2%(15)
Dim AC1%(15), AC2%(15)
Dim AVEN%(15), ALIS(15)
'
Dim REGTOPE&(15)
Dim REGACTI&(15)
Dim ISELTOP&(15)
Dim ISELACT&(15)
Dim FIDATIX#(15)
Dim YATASEL%
'
Dim SINLITEMP%
'
Dim TEBUS$(15), TEBUCASE$(15)
Dim TINIBU, YACARGA%
'
Dim CANRELI%
'
Sub RECARSQL(AAXX$)
   ' esto no hace nada ... se pone solo por compatibilidad con version SQL
End Sub
'
Private Sub ARTITUL_Change()
   If ARTITUL = "" Then
       SINI = "": LINI = ""
       If Visible = True Then Hide
     Else
       Call SEPARTITU(ARTITUL)  ' separa archivo y titulo
       Call OCULTABLAS          ' oculta todas las tablas
       Call IARTABLA            ' carga datos de tabla
   End If
End Sub

Private Sub BOTCANCE_Click()
    '
    If TASEL(ITABLAC%).Visible = True Then
       Text1.TEXT = ""
       TASEL(ITABLAC%).Visible = False
       If TASEL(ITABLAC%).ListCount > 0 Then
           TASEL(ITABLAC%).ListIndex = 0
       End If
       TABLA(ITABLAC%).Visible = True
       TASEL(ITABLAC%).Clear
       Exit Sub
       '
     Else
       '
       VD1(ITABLAC%).Caption = ""
       VD2(ITABLAC%).Caption = ""
       RENUAC(ITABLAC%).Caption = ""
       ARTITUL = ""
       On Error Resume Next
       SINI = "": LINI = ""
       Hide
       On Error GoTo 0
       Call FRESHALL
       '
   End If
   '
End Sub

Private Sub BOTCANCE_KeyDown(KeyCode As Integer, Shift As Integer)
   'Call REPLATEC(Keycode)
   'Call ENTRATEC(Keycode)
End Sub

Private Sub BOTFIL_Click()
   Call APERBASDAT("MASTER")
   If TIBASE$ = "DAT" Then
      Call MENSERR(24, "Filtros Complementarios no Disponibles\en la Presente Versión.")
      Exit Sub
   End If
   '
   If Combo1.ListCount < 1 Then
     For U& = 1 To ULTREG&("TATIMAT")
       XX$ = REGLEIDO$("TATIMAT", U&)
       TTXX$ = AJUSTI$(Mid$(XX$, 2), 120) + TRIM$(Str$(Asc(Left$(XX$, 1))))
       Combo1.AddItem TTXX$
     Next U&
   End If
   '
   If Combo2.ListCount < 1 Then
     For U& = 1 To ULTREG&("GRUCOVEN")
       XX$ = REGLEIDO$("GRUCOVEN", U&)
       TTXX$ = AJUSTI$(Mid$(XX$, 3), 120) + TRIM$(Str$(CVI(Left$(XX$, 2))))
       Combo2.AddItem TTXX$
     Next U&
   End If
   '
   If Combo3.ListCount < 1 Then
     For U& = 1 To ULTREG&("ACREDOR")
       XX$ = REGLEIDO$("ACREDOR", U&)
       TTXX$ = AJUSTI$(Mid$(XX$, 3), 120) + TRIM$(Str$(CVI(Left$(XX$, 2))))
       Combo3.AddItem TTXX$
     Next U&
   End If
   '
   Picture1.Top = Line2.Y1 + 10
   Height = Picture1.Top + Picture1.Height + 300
   Picture1.Visible = True
   TINIBU = Timer
   Timer1.Enabled = True
   '
End Sub

Private Sub Combo1_Click()
   Call FILTROSEC
End Sub

Private Sub Combo2_Click()
   Call FILTROSEC
End Sub

Private Sub FILTROSEC()
   '
   Call APERBASDAT("MASTER")
   If TIBASE$ = "DAT" Then Exit Sub
   '
   TINIBU = Timer
   TMT% = XVALO(Mid$(Combo1, 121))
   GRX% = XVALO(Mid$(Combo2, 121))
   PHB% = XVALO(Mid$(Combo3, 121))
   '
   Text1 = ""    ' PARA QUE SEA LA TABLA PRINCIPAL
   '
   SQLX$ = "SELECT * FROM Master "
   SQLX$ = SQLX$ + "WHERE STATUS >= 1 "
   SQLX$ = SQLX$ + "AND (" & TMT% & " = 0 OR TIMATE = " & TMT% & ") "
   SQLX$ = SQLX$ + "AND (" & GRX% & " = 0 OR GRUCOVEN = " & GRX% & ") "
   SQLX$ = SQLX$ + "AND (" & PHB% & " = 0 OR PROVHABI = " & PHB% & ") "
   SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
   ' Seteo el recordset
   Screen.MousePointer = 11
   Set MasterSel = ARTICULOS.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   Caption = "Maestro de Items de Stock"
   If (TMT% + GRX% + PHB%) < 1 Then
      TEBUS$(ITABLAC%) = LOADFILE$(LUDAT$ + "ECOMAST.SEL")
      TEBUCASE$(ITABLAC%) = UCase$(TEBUS$(ITABLAC%))
      GoTo 20
   End If
   '
   If TMT% > 0 Then Caption = Caption + " / " + TRIM$(ECOARCH$("TATIMAT", Chr$(TMT%)))
   If GRX% > 0 Then Caption = Caption + " / " + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRX%)))
   If PHB% > 0 Then Caption = Caption + " / " + TRIM$(ECOARCH$("ACREDOR", MKI$(PHB%)))
   '
   With MasterSel
      On Error Resume Next
      .MoveLast
      If Err Then
          FIN& = 0
        Else
          FIN& = .RecordCount
      End If
      TEBUS$(ITABLAC%) = Space$(64 * FIN&)
      TEBUCASE$(ITABLAC%) = Space$(64 * FIN&)
      '
      If FIN& > 0 Then
        .MoveFirst
        KU& = 1
        Do While (.EOF = False)
           TTXX$ = Space$(64)
           Call REPLA(TTXX$, !Codigo, 1, 16)
           Call REPLA(TTXX$, !Descripcion, 19, 46)
           TTYY$ = UCase$(TTXX$)
           Mid$(TEBUS$(ITABLAC%), KU&, 64) = TTXX$
           Mid$(TEBUCASE$(ITABLAC%), KU&, 64) = TTYY$
           KU& = KU& + 64
        .MoveNext
        Loop
      End If
   End With
   '
20 CAELEM& = Len(TEBUS$(ITABLAC%)) / AVEN(ITABLAC%)
   URETA& = ULTREG&(ARCHITA$)
   If URETA& < CAELEM& Then CAELEM& = URETA&
   '
   VScroll1(ITABLAC%).Min = 1
   MASCROLL& = CAELEM& - CANRELI%: If MASCROLL& < 1 Then MASCROLL& = 1
   OFSEVER& = 0
   If MASCROLL& > 32767 Then
      OFSEVER& = MASCROLL& - 32767
   End If
   VScroll1(ITABLAC%).Max = MASCROLL& - OFSEVER&
   VScroll1(ITABLAC%).Value = 1
   VScroll1(ITABLAC%).LargeChange = CANRELI% + 1
   '
   REGTOPE&(ITABLAC%) = 1
   REGACTI&(ITABLAC%) = 0
   '
   Call CARGATAB
   Call MUESTAB
   Screen.MousePointer = 1
   '
End Sub

Private Sub Combo3_Click()
   Call FILTROSEC
End Sub

Private Sub Combo1_Scroll()
   TINIBU = Timer
End Sub

Private Sub Combo2_Scroll()
   TINIBU = Timer
End Sub

Private Sub Combo3_Scroll()
   TINIBU = Timer
End Sub

Private Sub Command1_Click()
    Picture1.Visible = False
    Combo1.Visible = False
    Combo2.Visible = False
    Combo3.Visible = False
    Height = 3780
    Combo1.Visible = True
    Combo2.Visible = True
    Combo3.Visible = True
    Timer1.Enabled = False
End Sub

Private Sub Form_Activate()
   '
   If Picture1.Visible = True Then
       Picture1.Top = Line2.Y1 + 10
       Height = Picture1.Top + Picture1.Height + 300
       TINIBU = Timer
       Timer1.Enabled = True
     Else
       Height = 3780
   End If
   '
   If XVALO(LINI) > 0 And XVALO(LINI) < Screen.Width Then
       Left = XVALO(LINI)
       If Left > Screen.Width - Width Then
         Left = Screen.Width - Width
       End If
     Else
       Left = (Screen.Width - Width) / 2
   End If
   '
   V8700 = Screen.Height * 29 / 30
   If XVALO(SINI) > 0 And XVALO(SINI) < V8700 Then
       If XVALO(SINI) + Height > V8700 Then
           Top = XVALO(SINI) - Height - 440
         Else
           Top = XVALO(SINI)
       End If
     Else
       Top = (V8700 - Height) / 2
   End If
   '
   RECARTA% = 0
   If TABLA(ITABLAC%).ListCount < 1 Then
     Call CARTANUE
     RECARTA% = 1
   End If
   '
   ADAT$ = LUDAT$ + ARCHITA$ + ".DAT"
   ABUX$ = ARCHITA$
   If LOCARCH% = 1 Then
     ADAT$ = LUCOM$ + Mid$(ARCHITA$, 2) + ".WKF"
     ABUX$ = Mid$(ARCHITA$, 2)
   End If
   '
   On Error Resume Next
   FIDAT# = FileDateTime(ADAT$)
   On Error GoTo 0
   If FIDAT# <> FIDATIX#(ITABLAC%) Then
     Call CARTANUE
     RECARTA% = 1
   End If
   '
   If ABUSQUE = "" Then
     If RECARTA% > 0 Then
       If Text1 <> "" Then
         RETTT$ = Text1
         Text1 = ""
         Text1 = RETTT$
       End If
     End If
   End If
   '
   If ABUSQUE <> "" Then
      FILTER(ITABLAC%) = ""
      Text1 = ""
      TASEL(ITABLAC%).Visible = False
      TABLA(ITABLAC%).Visible = True
      Call BUSCARE
   End If
   Refresh
   '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Picture1.Visible = False Then
     Call REPLATEC(KeyCode)
     Call ENTRATEC(KeyCode)
     KeyCode = 0
   End If
  'KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Picture1.Visible = False Then
     KeyAscii = 0
   End If
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = 1
  VD1(ITABLAC%).Caption = ""
  VD2(ITABLAC%).Caption = ""
  RENUAC(ITABLAC%).Caption = ""
  ARTITUL = ""
  If TASEL(ITABLAC%).Visible = True Then
    ISELTOP&(ITABLAC%) = TASEL(ITABLAC%).TopIndex
    ISELACT&(ITABLAC%) = TASEL(ITABLAC%).ListIndex
  End If
  On Error Resume Next
  SINI = "": LINI = ""
  Hide
  On Error GoTo 0
End Sub

Private Sub Frame1_DblClick()
   Combo1.Clear
   YACARGA% = 0
   Call FILTROSEC
   Call BOTFIL_Click
End Sub

Private Sub Frame2_DblClick()
   Combo2.Clear
   YACARGA% = 0
   Call FILTROSEC
   Call BOTFIL_Click
End Sub

Private Sub Frame3_DblClick()
   Combo3.Clear
   YACARGA% = 0
   Call FILTROSEC
   Call BOTFIL_Click
End Sub

Private Sub Label1_DblClick()
   SINLITEMP% = 1
   Call Text1_Change
End Sub

Private Sub TABLA_Click(Index As Integer)
   REGACTI&(ITABLAC%) = REGTOPE&(ITABLAC%) + TABLA(ITABLAC%).ListIndex
   If REGTOPE&(ITABLAC%) >= VScroll1(ITABLAC%).Min Then
      On Error Resume Next
      VScroll1(ITABLAC%).Value = REGTOPE&(ITABLAC%) - OFSEVER&
      On Error GoTo 0
   End If
End Sub

Private Sub TABLA_DblClick(Index As Integer)
    TASELEC$ = TABLA(Index).TEXT
    VD1(Index) = TRIM$(Left$(TASELEC$, AC1%(Index)))
    VD2(Index) = TRIM$(Mid$(TASELEC$, 3 + AC1%(Index)))
    If ARCHITA$ <> "FECHA" And ARCHITA$ <> "SELFECHA" Then
      REGG& = TABLA(Index).ListIndex + REGTOPE&(Index)
      RENUAC(ITABLAC%).Caption = TRIM$(Str$(REGG&))
      Call SETREG(ARCHITA$, REGG&)
    End If
    SINI = "": LINI = ""
    Hide
    Call FRESHALL
End Sub
'

Private Sub TABLA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   '
   'Call REPLATECLA(Keycode)
   'Call ENTRATECLA(Keycode)
   '
End Sub

Private Sub TABLA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     SCMA = Screen.MousePointer
     YACT% = Int(Y / TEPRUEBA.Height)
     On Error Resume Next
     TABLA(Index).ListIndex = TABLA(Index).TopIndex + YACT%
     If Err Then Exit Sub
     On Error GoTo 0
     REGACTI(Index) = REGTOPE(Index) + YACT%
     ARGUBUS$ = TRIM$(Left$(TABLA(Index).TEXT, AC1%(Index)))
     If InStr(ARCHITA$, "ECOMAST") > 0 Or Left$(ARCHITA$, 5) = "LIPRE" Then
         Call MUESTRADATCOD(ARGUBUS$)
       ElseIf ARCHITA$ = "PROVED1" Or ARCHITA$ = "ACREDOR" Or ARCHITA$ = "PROVEPLA" Then
         Call MUESTRADATPRO(ARGUBUS$)
     End If
     Screen.MousePointer = SCMA
   End If
   '
End Sub

Private Sub TASEL_Click(Index As Integer)
    '
    Static TOPINANT
    ISELTOP&(Index) = TASEL(Index).TopIndex
    ISELACT&(Index) = TASEL(Index).ListIndex
    '
End Sub

Private Sub TASEL_DblClick(Index As Integer)
    TASELEC$ = TASEL(Index).TEXT
    VD1(Index) = TRIM$(Left$(TASELEC$, AC1%(Index)))
    VD2(Index) = TRIM$(Mid$(TASELEC$, 3 + AC1%(Index), AC2%(Index)))
    ISELTOP&(Index) = TASEL(Index).TopIndex
    ISELACT&(Index) = TASEL(Index).ListIndex
    ARCHU$ = ARCHITA$
    If ARCHU$ <> "FECHA" And ARCHU$ <> "SELFECHA" Then
      REGG& = XVALO(Mid$(TASELEC$, 3 + AC1%(Index) + AC2%(Index)))
      RENUAC(ITABLAC%).Caption = TRIM$(Str$(REGG&))
      Call SETREG(ARCHU$, REGG&)
    End If
    SINI = "": LINI = ""
    Hide
    Call FRESHALL
End Sub
'

Private Sub TASEL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   ABUSQUE = ""
   If Button = 2 Then
     YACT% = Int(Y / TEPRUEBA.Height)
     '
     On Error Resume Next
     TASEL(Index).ListIndex = TASEL(Index).TopIndex + YACT%
     If Err Then Exit Sub
     On Error GoTo 0
     '
     ARGUBUS$ = TRIM$(Left$(TASEL(Index).TEXT, AC1%(Index)))
     If InStr(ARCHITA$, "ECOMAST") > 0 Or Left$(ARCHITA$, 5) = "LIPRE" Then
         Call MUESTRADATCOD(ARGUBUS$)
       ElseIf ARCHITA$ = "PROVED1" Or ARCHITA$ = "ACREDOR" Or ARCHITA$ = "PROVEPLA" Then
         Call MUESTRADATPRO(ARGUBUS$)
     End If
   End If
   '
End Sub

Private Sub Text1_Change()
   '
   ABUS$ = UCase(Text1.TEXT)
   ABUSEC$ = ""
   PPXYZ% = InStr(ABUS$, "+")
   If PPXYZ% > 1 Then
      ABUSEC$ = TRIM$(Mid$(ABUS$, PPXYZ% + 1))
      ABUS$ = TRIM$(Left$(ABUS$, PPXYZ% - 1))
   End If
   '
   FILTER(ITABLAC%) = ABUS$
   If ABUS$ = "" Then
     If TASEL(ITABLAC%).Visible = True Then
       If TASEL(ITABLAC%).ListIndex >= 0 Then
         REATA& = XVALO(Mid$(TASEL(ITABLAC%).TEXT, 3 + AC1%(ITABLAC%) + AC2%(ITABLAC%)))
         If REGTOPE&(ITABLAC%) < REATA& - CANRELI% Or REGTOPE&(ITABLAC%) > REATA& Then REGTOPE&(ITABLAC%) = REATA& - 5
         If REGTOPE&(ITABLAC%) < 1 Then REGTOPE&(ITABLAC%) = 1
         Call CARGATAB
         LISCU& = TABLA(ITABLAC%).ListCount
         If REATA& > LISCU& + REGTOPE&(ITABLAC%) - 1 Then REATA& = LISCU& + REGTOPE&(ITABLAC%) - 1
         On Error Resume Next
         TABLA(ITABLAC%).ListIndex = REATA& - REGTOPE&(ITABLAC%)
         On Error GoTo 0
       End If
     End If
     TASEL(ITABLAC%).Visible = False
     TABLA(ITABLAC%).Visible = True
     ISELTOP&(ITABLAC%) = (-1)
     ISELACT&(ITABLAC%) = (-1)
     SINLITEMP% = 0
     Exit Sub
   End If
   '
   If SINLITEMP% > 0 Then Screen.MousePointer = 11
   '
   TASEL(ITABLAC%).ZOrder 0
   LOREBU% = AC1%(ITABLAC%) + AC2%(ITABLAC%) + 2
   TINI = Timer
   '
20 TASEL(ITABLAC%).Clear
   PINI& = 1
24 PPXY& = InStr(PINI&, TEBUCASE$(ITABLAC%), ABUS$)
   If PPXY& >= PINI& Then
      RETA& = Int(PPXY& / LOREBU%)
      TTTT$ = Mid$(TEBUS$(ITABLAC%), LOREBU% * RETA& + 1, LOREBU%) + Space$(6) + Str$(RETA& + 1)
      If ABUSEC$ = "" Or InStr(UCase$(TTTT$), ABUSEC$) > 0 Then
        TASEL(ITABLAC%).AddItem TTTT$
      End If
      PINI& = LOREBU% * RETA& + 1 + LOREBU%
      If Timer - TINI < 0.1 Or SINLITEMP% > 0 Then GoTo 24
      TASEL(ITABLAC%).AddItem " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ."
28 End If
   '
   TABLA(ITABLAC%).Visible = False
   TASEL(ITABLAC%).Visible = True
   If YATASEL% = 1 Then
     If ISELTOP&(ITABLAC%) >= 0 Then
       If ISELTOP&(ITABLAC%) < TASEL(ITABLAC%).ListCount Then
         TASEL(ITABLAC%).TopIndex = ISELTOP&(ITABLAC%)
       End If
     End If
     If ISELACT&(ITABLAC%) >= 0 Then
       If ISELACT&(ITABLAC%) < TASEL(ITABLAC%).ListCount Then
         TASEL(ITABLAC%).ListIndex = ISELACT&(ITABLAC%)
       End If
     End If
   End If
   '
   If TASEL(ITABLAC%).ListCount > 0 Then
     If TASEL(ITABLAC%).ListIndex < 0 Then
       TASEL(ITABLAC%).ListIndex = TASEL(ITABLAC%).TopIndex
     End If
   End If
   '
   If SINLITEMP% > 0 Then Screen.MousePointer = 1
   SINLITEMP% = 0
   '
End Sub

Private Sub Text1_DblClick()
    Text1.SelText = Text1.TEXT
End Sub

Private Sub Text1_GotFocus()
   KeyPreview = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 32 Then
      If KeyCode = 13 Or KeyCode >= 27 Then
        If KeyCode < 45 Then
          Call ENTRATEC(KeyCode)
        End If
      End If
    End If
End Sub

Sub REPLATEC(KeyCode)
    If KeyCode = 16 Then
      Text1.SetFocus
    End If
    '
    If KeyCode < 112 Then
      If KeyCode > 105 Then KeyCode = KeyCode - 64
      If KeyCode > 95 Then KeyCode = KeyCode - 48
    End If
End Sub

Sub SEPARTITU(ARTTX$)
   '
   PPXX% = InStr(ARTTX$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARTTX$)
   ARCHITA$ = TRIM$(UCase$(Left$(ARTTX$, PPXX% - 1)))
   If ARCHITA$ = "FECHA" Then ARCHITA = "SELFECHA"
   '
   If Left$(ARCHITA$, 1) = "^" Then
       ARCHITA$ = Mid$(ARCHITA$, 2)
     ElseIf Left$(ARCHITA$, 1) = "*" Then
       FRXX$ = Mid$(ARCHITA$, 2) + ".PRF"
       'If VERDEMO% = 1 Then
       '  FRXX$ = "\FX_PROFILS\" + FRXX$
       'End If
       ARCHITA$ = TRIM$(Mid$(LOADFILE$(FRXX$), 57, 8))
   End If
   '
   TITULTA$ = Mid$(ARTTX$, PPXX% + 1)
   If TRIM$(TITULTA$) = "" Then
      TITULTA$ = TRIM$(TITUARCH$(ARCHITA$))
   End If
   Caption = TITULTA$
   '
   If ARCHITA$ = "" Or Len(ARCHITA$) > 9 Then
      Call MENSERR(24, "Error de Invocación de Rutina 'SELECHO' - Consulte Soporte de Sistemas.")
      ARCHITA$ = ""
      SINI = "": LINI = ""
      Hide
      Exit Sub
   End If
   '
   If ARCHITA$ = "MASTER" Then
       ARCHITA$ = "ECOMAST"
     ElseIf ARCHITA$ = "FECHA" Or ARCHITA$ = "SELFECHA" Then
       ARCHITA$ = "SELFECHA"
   End If
   '
   If Left$(ARCHITA$, 1) = "*" Then
        FRXX$ = Mid$(ARCHITA$, 2) + ".PRF"
        'If VERDEMO% = 1 Then
        '  FRXX$ = "\FX_PROFILS\" + FRXX$
        'End If
        Dim YYY As String * 128
        NKX% = FreeFile
        On Error Resume Next
        Open FRXX$ For Random Access Read Shared As #NKX% Len = 128
        Get #NKX%, 1, YYY$
        ARCHITA$ = Mid$(YYY$, 57, 8)
        Close #NKX%: NKX% = 0
        On Error GoTo 0
      ElseIf Left$(ARCHITA$, 1) = "^" Then
        ARCHITA$ = Mid$(ARCHITA$, 2)
   End If
   '
End Sub

Sub OCULTABLAS()
   '
   For KKII% = 0 To 15
      MARCOTA(KKII%).Visible = False
   Next KKII%
   '
End Sub

Sub IARTABLA()
   '
   ADAT$ = LUDAT$ + ARCHITA$ + ".DAT"
   ABUX$ = ARCHITA$
   If LOCARCH% = 1 Then
     ADAT$ = LUCOM$ + Mid$(ARCHITA$, 2) + ".WKF"
     ABUX$ = Mid$(ARCHITA$, 2)
   End If
   '
   On Error Resume Next
   FIDAT# = FileDateTime(ADAT$)
   On Error GoTo 0
   '
   ARCHU16$ = AJUSTI$("@@" + ARCHITA$, 16)
   PPXX% = InStr(VECARTA$, ARCHU16$)
   If (PPXX% Mod 16) = 1 Then
      ITABLAC% = (PPXX% - 1) / 16
      If TABLA(ITABLAC%).ListCount < 1 Or FIDAT# <> FIDATIX#(ITABLAC%) Then
        Call CARDATAB
        Call CARTANUE
      End If
      Call MUESTAB
      Exit Sub
   End If
   '
   If Len(VECARTA$) >= 256 Then
      Call MENSERR(24, "Demasiadas Tablas Abiertas.\Consulte al Soporte de Sistemas.")
      Call FINAL("")
   End If
   '
   VECARTA$ = VECARTA$ + ARCHU16$
   ITABLAC% = Len(VECARTA$) / 16 - 1
   ATABLA(ITABLAC%).Caption = ARCHITA$
   VD1(ITABLAC%).Caption = ""
   VD2(ITABLAC%).Caption = ""
   RENUAC(ITABLAC%).Caption = ""
   '
10 If ARCHITA$ = "SELFECHA" Or ARCHITA$ = "FECHA" Then
       Call GETAFECH
   End If
   Call CARDATAB
   Call CARTANUE
   Call MUESTAB
   FIDATIX#(ITABLAC%) = FIDAT#
   '
   If ABUSQUE <> "" Then
      FILTER(ITABLAC%) = ""
      Text1 = ""
      TASEL(ITABLAC%).Visible = False
      TABLA(ITABLAC%).Visible = True
      Call BUSCARE
   End If
   '
   TASEL(ITABLAC%).Top = 0
   TASEL(ITABLAC%).Left = 0
   TASEL(ITABLAC%).Height = Text1.Top
   TASEL(ITABLAC%).Width = ALIS(ITABLAC%) + 300
   TASEL(ITABLAC%).BackColor = TASEL(0).BackColor
   '
End Sub

Sub CARDATAB()
   '
   TV1%(ITABLAC%) = TIPFIELD%(ARCHITA$, 1): TV2%(ITABLAC%) = TIPFIELD%(ARCHITA$, 2)
   LF1%(ITABLAC%) = LENFIELD%(ARCHITA$, 1): LF2%(ITABLAC%) = LENFIELD%(ARCHITA$, 2)
   PR1%(ITABLAC%) = POSFIELD%(ARCHITA$, 1): PR2%(ITABLAC%) = POSFIELD%(ARCHITA$, 2)
   LD1%(ITABLAC%) = DECFIELD%(ARCHITA$, 1): LD2%(ITABLAC%) = DECFIELD%(ARCHITA$, 2)
   '
   AC1%(ITABLAC%) = LF1%(ITABLAC%)
   JT1%(ITABLAC%) = 1
   'LD1%(ITABLAC%) = 0
   If TV1%(ITABLAC%) = 1 Or TV1%(ITABLAC%) = 2 Then
        JT1%(ITABLAC%) = 2
        If LF1%(ITABLAC%) < 2 Then AC1%(ITABLAC%) = 3 Else AC1%(ITABLAC%) = 5
      ElseIf TV1%(ITABLAC%) = 3 Or TV1%(ITABLAC%) = 4 Then
        JT1%(ITABLAC%) = 2  ': LD1%(ITABLAC%) = 1
        If LF1%(ITABLAC%) < 8 Then AC1%(ITABLAC%) = 7 Else AC1%(ITABLAC%) = 12
        If LD1%(ITABLAC%) > 0 Then AC1%(ITABLAC%) = AC1%(ITABLAC%) + 2
      ElseIf TV1%(ITABLAC%) = 5 Then
        AC1%(ITABLAC%) = 8
      ElseIf TV1%(ITABLAC%) = 7 Then
        AC1%(ITABLAC%) = 16
   End If
   '
   AC2%(ITABLAC%) = LF2%(ITABLAC%)
   JT2%(ITABLAC%) = 1
   'LD2%(ITABLAC%) = 0
   If TV2%(ITABLAC%) = 1 Or TV2%(ITABLAC%) = 2 Then
        JT2%(ITABLAC%) = 2
        If LF2%(ITABLAC%) < 2 Then AC2%(ITABLAC%) = 3 Else AC2%(ITABLAC%) = 5
      ElseIf TV2%(ITABLAC%) = 3 Or TV2%(ITABLAC%) = 4 Then
        JT2%(ITABLAC%) = 2
        'LD2%(ITABLAC%) = 2
        If LF2%(ITABLAC%) < 8 Then AC2%(ITABLAC%) = 7 Else AC2%(ITABLAC%) = 12
        If LD2%(ITABLAC%) > 0 Then AC2%(ITABLAC%) = AC2%(ITABLAC%) + 2
      ElseIf TV2%(ITABLAC%) = 5 Then
        AC2%(ITABLAC%) = 8
      ElseIf TV2%(ITABLAC%) = 7 Then
        AC2%(ITABLAC%) = 16
    End If
    '
    If AC1%(ITABLAC%) + AC2%(ITABLAC%) > 70 Then AC2%(ITABLAC%) = 70 - AC1%(ITABLAC%)
    AVEN%(ITABLAC%) = AC1%(ITABLAC%) + 2 + AC2%(ITABLAC%)
    If AVEN%(ITABLAC%) > 72 Then AVEN%(ITABLAC%) = 72
    TEPRUEBA.Caption = String$(2 + AVEN%(ITABLAC%), "*")
    ALIS(ITABLAC%) = TEPRUEBA.Width
    TABLA(ITABLAC%).Width = ALIS(ITABLAC%)
    '
    TABLA(ITABLAC%).Top = 0
    TABLA(ITABLAC%).Left = 0
    '
    Line1(ITABLAC%).X1 = TABLA(ITABLAC%).Width + 20
    Line1(ITABLAC%).X2 = Line1(ITABLAC%).X1
    Line1(ITABLAC%).Y1 = 0
    Line1(ITABLAC%).Y2 = TABLA(ITABLAC%).Height
    '
    VScroll1(ITABLAC%).Left = Line1(ITABLAC%).X1 + 40
    VScroll1(ITABLAC%).Top = 0
    VScroll1(ITABLAC%).Height = TABLA(ITABLAC%).Height
    '
    MARCOTA(ITABLAC%).Width = VScroll1(ITABLAC%).Left + VScroll1(ITABLAC%).Width + 50
    TASEL(ITABLAC%).Top = 0
    TASEL(ITABLAC%).Left = 0
    TASEL(ITABLAC%).Height = TABLA(ITABLAC%).Height
    TASEL(ITABLAC%).Width = MARCOTA(ITABLAC%).Width
    '
End Sub

Sub MUESTAB()
   '
   Line2.X2 = MARCOTA(ITABLAC%).Width
   Width = MARCOTA(ITABLAC%).Width + 60
   If XVALO(LINI) > 0 And XVALO(LINI) < Screen.Width Then
       Left = XVALO(LINI)
       If Left > Screen.Width - Width Then
         Left = Screen.Width - Width
       End If
     Else
       Left = (Screen.Width - Width) / 2
   End If
   '
   V8700 = Screen.Height * 29 / 30
   If XVALO(SINI) > 0 And XVALO(SINI) < V8700 Then
       If XVALO(SINI) + Height > V8700 Then
           Top = XVALO(SINI) - Height - 440
         Else
           Top = XVALO(SINI)
       End If
     Else
       Top = (V8700 - Height) / 2
   End If
  
   BOTERA.Width = 430
   BOTINSE.Width = 430
   BOTERA.Visible = True
   BOTINSE.Visible = True
   '
   BOTCANCE.Left = ALIS(ITABLAC%) - BOTCANCE.Width
   '
   If ALIS(ITABLAC%) < 3000 Then
       BOTERA.Width = 0
       BOTINSE.Width = 0
       BOTERA.Visible = False
       BOTINSE.Visible = False
   End If
   BOTERA.Left = BOTCANCE.Left - BOTERA.Width
   BOTINSE.Left = BOTERA.Left - BOTINSE.Width
   BOTINSE.Enabled = False
   BOTERA.Enabled = False
   '
   BOFILTRA% = 0
   BOTFIL.Visible = False
   If ARCHITA$ = "ECOMAST" Then
      BOTFIL.Visible = True
      BOFILTRA% = 1
      BOTFIL.Left = BOTINSE.Left - BOTFIL.Width
   End If
   '
   On Error Resume Next
   Text1.Width = BOTINSE.Left - Text1.Left - 100
   If BOFILTRA% = 1 Then
      Text1.Width = Text1.Width - BOTFIL.Width
   End If
   On Error GoTo 0
   MARCOTA(ITABLAC%).Visible = True
   '
   If ABUSQUE <> "" Then
      YATASEL% = 0
      FILTER(ITABLAC%) = ""
      Call BUSCARE
      Refresh
      Exit Sub
   End If
   '
   YATASEL% = 0
   If FILTER(ITABLAC%) <> "" Then YATASEL% = 1
   Text1 = FILTER(ITABLAC%)
   YATASEL% = 0
   If Text1 = "" Then
     TASEL(ITABLAC%).Visible = False
     TABLA(ITABLAC%).Visible = True
   End If
   Refresh
   '
End Sub

Sub BUSCARE()
   '
   ABUS$ = UCase$(ABUSQUE)
   If JT1%(ITABLAC%) = 2 Then
     ABUS$ = AJUSTD$(ABUS$, AC1%(ITABLAC%))
   End If
   LOREBU% = AC1%(ITABLAC%) + AC2%(ITABLAC%) + 2
   '
   PINI& = 1
24 PPXY& = InStr(PINI&, TEBUCASE$(ITABLAC%), ABUS$)
   If PPXY& < PINI& Then Exit Sub
   If PPXY& Mod LOREBU% <> 1 Then
     PINI& = PPXY& + 1
     GoTo 24
   End If
   RETA& = 1 + Int(PPXY& / LOREBU%)
   REGACTI&(ITABLAC%) = RETA&
   REGTOPE&(ITABLAC%) = RETA& - 5
   If REGTOPE&(ITABLAC%) < 1 Then REGTOPE&(ITABLAC%) = 1
   Call CARGATAB
   '
End Sub
'
Sub CARTANUE()
   '
   VUECARTA% = 0
10 Call RECARSEL
   '
   YAMU% = 0
   LOCARCH% = 0: If Left$(ARCHITA$, 1) = "!" Then LOCARCH% = 1
   ADAT$ = LUDAT$ + ARCHITA$
   If LOCARCH% = 1 Then
     ADAT$ = LUCOM$ + Mid$(ARCHITA$, 2)
   End If
   '
   TEBUS$(ITABLAC%) = LOADFILE$(ADAT$ + ".SEL")
   If VUECARTA% < 1 Then
     If TEBUS$(ITABLAC%) = "" Then
       If ULTREG&(ARCHITA$) > 0 Then
         XCX$ = REGLEIDO$(ARCHITA$, 1)
         Call GRAREG(ARCHITA$, XCX$, 1)
         Call CIERRARCH(ARCHITA$)
         VUECARTA% = 1
         GoTo 10
       End If
     End If
   End If
   '
   TINI = Timer
   TEBUCASE$(ITABLAC%) = UCase$(TEBUS$(ITABLAC%))
   '
   TABLA(ITABLAC%).Clear
   TABLA(ITABLAC%).Visible = True
   CAELEM& = Len(TEBUS$(ITABLAC%)) / AVEN(ITABLAC%)
   URETA& = ULTREG&(ARCHITA$)
   If URETA& < CAELEM& Then CAELEM& = URETA&
   '
   VScroll1(ITABLAC%).Min = 1
   MASCROLL& = CAELEM& - CANRELI%: If MASCROLL& < 1 Then MASCROLL& = 1
   OFSEVER& = 0
   If MASCROLL& > 32767 Then
      OFSEVER& = MASCROLL& - 32767
   End If
   VScroll1(ITABLAC%).Max = MASCROLL& - OFSEVER&
   VScroll1(ITABLAC%).Value = 1
   VScroll1(ITABLAC%).LargeChange = CANRELI% + 1
   '
   REGTOPE&(ITABLAC%) = 1
   REGACTI&(ITABLAC%) = 0
   If ARCHITA$ = "SELFECHA" Or ARCHITA$ = "FECHA" Then
     REGACTI&(ITABLAC%) = REFECHAC&
     REGTOPE&(ITABLAC%) = REGACTI&(ITABLAC%) - 5
   End If
   Call CARGATAB
   If ARCHITA$ = "ECOMAST" Then
      Call FILTROSEC
   End If
   '
90 MARCOTA(ITABLAC%).Visible = True
   TABLA(ITABLAC%).Refresh
   Screen.MousePointer = 1
   '
End Sub

Sub RECARSEL()   ' Re-Genera archivo .SEL
   '
   Static ACTUSEL$
   Static ACSELOC$
   '
   If ACTUSEL$ = "" Or ACSELOC$ = "" Then
     ACTUSEL$ = LOADFILE$(LUDAT$ + "ACTUSEL.DRV")
     ACSELOC$ = LOADFILE$(LUCOM$ + "ACTUSEL.DRV")
   End If
   '
   LOCARCH% = 0: If Left$(ARCHITA$, 1) = "!" Then LOCARCH% = 1
   ADAT$ = LUDAT$ + ARCHITA$ + ".DAT"
   ABUX$ = ARCHITA$
   ADASEL$ = LUDAT$ + ARCHITA$ + ".SEL"
   If LOCARCH% = 1 Then
     ADAT$ = LUCOM$ + Mid$(ARCHITA$, 2) + ".WKF"
     ABUX$ = Mid$(ARCHITA$, 2)
     ADASEL$ = LUCOM$ + ARCHITA$ + ".SEL"
   End If
   '
   On Error Resume Next
   FIDAT# = FileDateTime(ADAT$)
   On Error GoTo 0
   FIDATIX#(ITABLAC%) = FIDAT#
   '
   ARCHU10$ = AJUSTI$("@@" + ABUX$, 10)
   '
   If LOCARCH% = 0 Then
       ISELAR$ = ACTUSEL$
       PPXX% = InStr(ACTUSEL$, ARCHU10$)
       If (PPXX% Mod 18) = 1 Then GoTo 10
       While (Len(ACTUSEL$) Mod 18) <> 0
         ACTUSEL$ = ACTUSEL$ + " "
       Wend
       PPXX% = 1 + Len(ACTUSEL$)
       ACTUSEL$ = ACTUSEL$ + ARCHU10$ + MKD$(0)
       ISELAR$ = ACTUSEL$
     Else
       ISELAR$ = ACSELOC$
       PPXX% = InStr(ACSELOC$, ARCHU10$)
       If (PPXX% Mod 18) = 1 Then GoTo 10
       While (Len(ACSELOC$) Mod 18) <> 0
         ACSELOC$ = ACSELOC$ + " "
       Wend
       PPXX% = 1 + Len(ACSELOC$)
       ACSELOC$ = ACSELOC$ + ARCHU10$ + MKD$(0)
       ISELAR$ = ACSELOC$
   End If
   '
10 If EXISTE%(ADASEL$) > 0 Then
     If FIDAT# = CVD(Mid$(ISELAR$, PPXX% + 10, 8)) Then
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   VELEC$ = LOADFILE$(ADAT$)
   LOFI% = RECLEN%(ARCHITA$)
   ALISTA% = AC1%(ITABLAC%) + AC2%(ITABLAC%) + 2
   URE& = ULTREG&(ARCHITA$)
   '
   TEBUS$(ITABLAC%) = Space$(URE& * ALISTA%)
   POSIN& = 1
   '
   For U& = 1 To Len(VELEC$) Step LOFI%
     If U& <= URE& * LOFI% Then
       XXX$ = Mid$(VELEC$, U&, LOFI)
       XXZ$ = Left$(XXX$, LF1%(ITABLAC%))
          If ADEUDO% = 1 Then If CVI(XXZ$) < 1 Then GoTo 19
       TX1$ = CSTRING$(XXZ$, TV1%(ITABLAC%), AC1%(ITABLAC%), LD1%(ITABLAC%), JT1%(ITABLAC%))
       XXZ$ = Mid$(XXX$, PR2%(ITABLAC%), LF2%(ITABLAC%))
       TX2$ = CSTRING$(XXZ$, TV2%(ITABLAC%), AC2%(ITABLAC%), LD2%(ITABLAC%), JT2%(ITABLAC%))
       Mid$(TEBUS$(ITABLAC%), POSIN&, ALISTA%) = TX1$ + "  " + TX2$
       POSIN& = POSIN& + ALISTA%
     End If
19 Next U&
   '
   If LOCARCH% = 0 Then
       Call SAVEFILE(LUDAT$ + ARCHITA$ + ".SEL", TEBUS$(ITABLAC%))
       Call REPLA(ACTUSEL$, MKD$(FIDAT#), PPXX% + 10, 8)
       Call SAVEFILE(LUDAT$ + "ACTUSEL.DRV", ACTUSEL$)
     Else
       Call SAVEFILE(LUCOM$ + ABUX$ + ".SEL", TEBUS$(ITABLAC%))
       Call REPLA(ACSELOC$, MKD$(FIDAT#), PPXX% + 10, 8)
       Call SAVEFILE(LUCOM$ + "ACTUSEL.DRV", ACSELOC$)
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARGATAB()
   '
   REGINI& = REGTOPE&(ITABLAC%)
   If REGINI& < 1 Then REGINI& = 1
   If REGACTI&(ITABLAC%) < 1 Then REGACTI&(ITABLAC%) = 1
   TABLA(ITABLAC%).Clear
   '
   If CANRELI% < 1 Then
      TEPRUEBA.Font = TABLA(ITABLAC%).Font
      TEPRUEBA.FontName = TABLA(ITABLAC%).FontName
      TEPRUEBA.FontSize = TABLA(ITABLAC%).FontSize
      TEPRUEBA.FontBold = TABLA(ITABLAC%).FontBold
      TEPRUEBA.Caption = "ABCDE"
      CANRELI% = TABLA(ITABLAC%).Height / TEPRUEBA.Height - 1
      If CANRELI% < 1 Then CANRELI% = 11
   End If
   '
   For U& = REGINI& To REGINI& + CANRELI%
     POLEC& = (U& - 1) * AVEN(ITABLAC%) + 1
     If POLEC& < Len(TEBUS$(ITABLAC%)) Then
       TABLA(ITABLAC%).AddItem Mid$(TEBUS$(ITABLAC%), POLEC&, AVEN%(ITABLAC%))
     End If
   Next U&
   '
   If REGACTI&(ITABLAC%) >= REGINI& Then
     If REGACTI&(ITABLAC%) < REGINI& + TABLA(ITABLAC%).ListCount + 0 Then
        REATA& = REGACTI&(ITABLAC%) - REGINI&
        If REATA& > TABLA(ITABLAC%).ListCount - 1 Then REATA& = TABLA(ITABLAC%).ListCount - 1
        TABLA(ITABLAC%).ListIndex = REATA&
     End If
   End If
   TABLA(ITABLAC%).Refresh
   '
   If TABLA(ITABLAC%).ListCount > 0 Then
     If REGTOPE&(ITABLAC%) < REGINI& Then
       REGTOPE&(ITABLAC%) = REGINI&
     End If
   End If
   '
End Sub
'

Private Sub text1_lostfocus()
   KeyPreview = True
End Sub

Private Sub Timer1_Timer()
   If Timer > TINIBU + 15 Then
      Call Command1_Click
   End If
End Sub

Private Sub VScroll1_Change(Index As Integer)
   If VScroll1(ITABLAC%).Value + OFSEVER& <> REGTOPE&(ITABLAC%) Then
     REGTOPE&(ITABLAC%) = VScroll1(ITABLAC%).Value + OFSEVER&
     Call CARGATAB
   End If
End Sub

Private Sub VScroll1_Scroll(Index As Integer)
   If CONTROBARRA% >= 0 Then
     If TABLA(ITABLAC%).Visible = True Then
         REGTOPE&(ITABLAC%) = VScroll1(ITABLAC%).Value + OFSEVER&
       ElseIf TASEL(ITABLAC%).Visible = True Then
         TASEL(ITABLAC%).TopIndex = VScroll1(ITABLAC%).Value + OFSEVER& - 1
     End If
   End If
   '
   If TABLA(ITABLAC%).Visible = True Then
     Call CARGATAB
   End If

End Sub

Sub ENTRATEC(KeyCode)
   '
   If KeyCode = 27 Then
      If TRIM$(Text1.TEXT) <> "" Then
         Text1.TEXT = ""
       Else
         VD1(ITABLAC%) = ""
         VD2(ITABLAC%) = ""
         RENUAC(ITABLAC%).Caption = ""
         SINI = "": LINI = ""
         Hide
      End If
      Exit Sub
      '
    ElseIf KeyCode = 8 Then
      POMOUSE = Text1.SelStart
      If Text1.TEXT <> "" Then
        Text1.TEXT = Left$(Text1.TEXT, Len(Text1.TEXT) - 1)
      End If
      If POMOUSE > 1 Then Text1.SelStart = POMOUSE
      Exit Sub
      '
    ElseIf KeyCode = 37 Or KeyCode = 39 Then
      Text1.SetFocus
      Exit Sub
      '
    ElseIf KeyCode >= 42 Or KeyCode = 32 Then
      If KeyCode < 120 Then
        Text1.TEXT = Text1.TEXT + Chr$(KeyCode)
        Text1.SelStart = Len(Text1.TEXT)
        Text1.SetFocus
        Exit Sub
      End If
   End If
   '
   If TASEL(ITABLAC%).Visible = True Then
      If KeyCode = 13 Then
        If TASEL(ITABLAC%).ListIndex >= 0 Then
          VD1(ITABLAC%).Caption = Left$(TASEL(ITABLAC%).TEXT, AC1%(ITABLAC%))
          VD2(ITABLAC%).Caption = Mid$(TASEL(ITABLAC%).TEXT, 3 + AC1%(ITABLAC%), AC2%(ITABLAC%))
          REGG& = XVALO(Mid$(TASEL(ITABLAC%).TEXT, 3 + AC1%(ITABLAC%) + AC2%(ITABLAC%)))
          RENUAC(ITABLAC%).Caption = TRIM$(Str$(REGG&))
          Call SETREG(ARCHITA$, REGG&)
          RETA& = REGG&
          SINI = "": LINI = ""
          Hide
          Exit Sub
        End If
       ElseIf KeyCode = 38 Then
        If TASEL(ITABLAC%).ListIndex > 0 Then
          TASEL(ITABLAC%).ListIndex = TASEL(ITABLAC%).ListIndex - 1
        End If
       ElseIf KeyCode = 40 Then
        If TASEL(ITABLAC%).ListIndex < TASEL(ITABLAC%).ListCount - 1 Then
          TASEL(ITABLAC%).ListIndex = TASEL(ITABLAC%).ListIndex + 1
        End If
       ElseIf KeyCode = 34 Then
        If TASEL(ITABLAC%).ListIndex < TASEL(ITABLAC%).ListCount - 1 Then
          LINDENU& = TASEL(ITABLAC%).ListIndex + (CANRELI% + 1)
          If LINDENU& > TASEL(ITABLAC%).ListCount - 1 Then LINDENU& = TASEL(ITABLAC%).ListCount - 1
          TASEL(ITABLAC%).ListIndex = LINDENU&
        End If
       ElseIf KeyCode = 33 Then
        If TASEL(ITABLAC%).ListIndex > 0 Then
          LINDENU& = TASEL(ITABLAC%).ListIndex - (CANRELI% + 1)
          If LINDENU& < 0 Then LINDENU& = 0
          TASEL(ITABLAC%).ListIndex = LINDENU&
        End If
       ElseIf KeyCode = 36 Then
        On Error Resume Next
        TASEL(ITABLAC%).ListIndex = 0
        On Error GoTo 0
       ElseIf KeyCode = 35 Then
        TASEL(ITABLAC%).ListIndex = TASEL(ITABLAC%).ListCount - 1
      End If
      TASEL(ITABLAC%).SetFocus
      Exit Sub
   End If
   '
   If TABLA(ITABLAC%).Visible = True Then
     TABLA(ITABLAC%).SetFocus
     If KeyCode = 13 Then
        If TABLA(ITABLAC%).ListIndex >= 0 Then
          VD1(ITABLAC%) = Left$(TABLA(ITABLAC%).TEXT, AC1%(ITABLAC%))
          VD2(ITABLAC%) = TRIM$(Mid$(TABLA(ITABLAC%).TEXT, 3 + AC1%(ITABLAC%), AC2%(ITABLAC%)))
          REGG& = TABLA(ITABLAC%).ListIndex + REGTOPE&(ITABLAC%)
          RENUAC(ITABLAC%).Caption = TRIM$(Str$(REGG&))
          Call SETREG(ARCHITA$, REGG&)
          SINI = "": LINI = ""
          Hide
        End If
        Exit Sub
     ElseIf KeyCode = 38 Then
       If TABLA(ITABLAC%).ListIndex < 0 Then TABLA(ITABLAC%).ListIndex = 0
       If TABLA(ITABLAC%).ListIndex = 0 Then
           If REGACTI&(ITABLAC%) > 1 Then
             REGTOPE(ITABLAC%) = REGTOPE(ITABLAC%) - 1
             Call CARGATAB
             TABLA(ITABLAC%).ListIndex = 0
           End If
         Else
           TABLA(ITABLAC%).ListIndex = TABLA(ITABLAC%).ListIndex - 1
       End If
     ElseIf KeyCode = 40 Then
       If TABLA(ITABLAC%).ListIndex = CANRELI% Then
           If REGTOPE(ITABLAC%) < VScroll1(ITABLAC%).Max + OFSEVER& Then
             REGTOPE(ITABLAC%) = REGTOPE(ITABLAC%) + 1
             Call CARGATAB
             TABLA(ITABLAC%).ListIndex = CANRELI%
           End If
         Else
           If TABLA(ITABLAC%).ListIndex < TABLA(ITABLAC%).ListCount - 1 Then
             TABLA(ITABLAC%).ListIndex = TABLA(ITABLAC%).ListIndex + 1
           End If
       End If
     ElseIf KeyCode = 33 Then
       REGACTI(ITABLAC%) = REGACTI(ITABLAC%) - (CANRELI% + 1)
       REGTOPE(ITABLAC%) = REGTOPE(ITABLAC%) - (CANRELI% + 1)
       If REGTOPE(ITABLAC%) < 1 Then REGTOPE(ITABLAC%) = 1
       If REGACTI(ITABLAC%) < REGTOPE(ITABLAC%) Then REGACTI(ITABLAC%) = REGTOPE(ITABLAC%)
       Call CARGATAB
     ElseIf KeyCode = 34 Then
       REGACTI(ITABLAC%) = REGACTI(ITABLAC%) + (CANRELI% + 1)
       REGTOPE(ITABLAC%) = REGTOPE(ITABLAC%) + (CANRELI% + 1)
       If REGACTI(ITABLAC%) > ULTREG&(ARCHITA$) Then REGACTI(ITABLAC%) = ULTREG&(ARCHITA$)
       If REGTOPE(ITABLAC%) > VScroll1(ITABLAC%).Max + OFSEVER& Then REGTOPE&(ITABLAC%) = VScroll1(ITABLAC%).Max + OFSEVER&
       Call CARGATAB
     ElseIf KeyCode = 36 Then
       REGACTI(ITABLAC%) = 1
       REGTOPE(ITABLAC%) = 1
       Call CARGATAB
     ElseIf KeyCode = 35 Then
       REGACTI(ITABLAC%) = ULTREG&(ARCHITA$)  ' FORSELECHO.SCROLLBA.Max
       REGTOPE(ITABLAC%) = REGACTI(ITABLAC%) - CANRELI%
       If REGTOPE(ITABLAC%) < 1 Then REGTOPE(ITABLAC%) = 1
       Call CARGATAB
    End If
    '
  End If
  '
End Sub

Sub GETAFECH()
  '
  DIAC& = Int(Now)
  DITAFE& = Int(FIDATIM#("SELFECHA"))
  '
  UREFE& = ULTREG&("SELFECHA")
  If DITAFE& <> DIAC& Or UREFE& < 365 Or UREFE& > 750 Then
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    '
    On Error Resume Next
    Kill LUDAT$ + "SELFECHA.DAT"
    On Error GoTo 0
    '
    HOOI% = HOY(HO$)
    ANX = XVALO(Mid$(HO$, 7, 2)) - 1: If ANX < 0 Then ANX = ANX + 100
    ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
    DESDE$ = Left$(HO$, 6) + ANXS$
    ANX = XVALO(Mid$(HO$, 7, 2)) + 1: If ANX > 99 Then ANX = ANX - 100
    ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
    HASTA$ = Left$(HO$, 6) + ANXS$
    '
    JJ& = 0
    HAS = FECHANUM(HASTA$)
    FEX = FECHANUM(DESDE$)
    While FEX <= HAS
      TX1$ = FECHATEX$(FEX)
      TX2$ = DSEM$(FEX)
      TX12$ = TX1$ + "  " + TX2$
      If TRIM$(TX12$) <> "" Then
        JJ& = JJ& + 1
        TTXX$ = MKI$(Int(FEX)) + TX2$
        Call GRAREG("SELFECHA", TTXX$, JJ&)
      End If
      FEY = DIPOST(FEX)
      FEX = FEY
    Wend
    Call SETULTREG("SELFECHA", JJ&)
    Call CIERRARCH("SELFECHA")
    Screen.MousePointer = SCMA
  End If
  '
  REFECHAC& = ULTREG&("SELFECHA") / 2
9 If CVI(Left$(REGLEIDO$("SELFECHA", REFECHAC&), 2)) < HOOI% Then
    REFECHAC& = REFECHAC& + 1
    GoTo 9
  End If
  '
End Sub
