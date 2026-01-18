VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SOLINT07 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Interna de Materiales"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   8130
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1380
      Left            =   10395
      Picture         =   "SOLINT07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Aprueba, Graba e Imprime Solicitud de Materiales"
      Top             =   6615
      Width           =   800
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1575
      Left            =   7035
      TabIndex        =   99
      Top             =   5985
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2778
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   285
      Left            =   9990
      TabIndex        =   3
      Top             =   1320
      Width           =   195
   End
   Begin VB.TextBox Text6 
      Height          =   280
      Left            =   1230
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox Picture6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   150
      ScaleHeight     =   3465
      ScaleWidth      =   10035
      TabIndex        =   91
      Top             =   2280
      Width           =   10095
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFF0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6225
         Left            =   0
         ScaleHeight     =   6225
         ScaleWidth      =   9855
         TabIndex        =   112
         Top             =   0
         Width           =   9850
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   9
            Left            =   9525
            TabIndex        =   151
            Top             =   5835
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   8
            Left            =   9525
            TabIndex        =   148
            Top             =   5235
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   7
            Left            =   9525
            TabIndex        =   145
            Top             =   4640
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   6
            Left            =   9525
            TabIndex        =   142
            Top             =   4030
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   5
            Left            =   9525
            TabIndex        =   139
            Top             =   3435
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   4
            Left            =   9525
            TabIndex        =   136
            Top             =   2850
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   3
            Left            =   9525
            TabIndex        =   133
            Top             =   2220
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   1
            Left            =   9525
            TabIndex        =   130
            Top             =   1035
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   2
            Left            =   9525
            TabIndex        =   127
            Top             =   1635
            Width           =   195
         End
         Begin VB.CommandButton Command12 
            Caption         =   "."
            Height          =   210
            Index           =   0
            Left            =   9520
            TabIndex        =   124
            Top             =   400
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   9
            Left            =   90
            TabIndex        =   122
            Top             =   5550
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   8
            Left            =   90
            TabIndex        =   121
            Top             =   4950
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   7
            Left            =   90
            TabIndex        =   120
            Top             =   4350
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   6
            Left            =   90
            TabIndex        =   119
            Top             =   3750
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   118
            Top             =   3150
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   117
            Top             =   2550
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   116
            Top             =   1950
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   115
            Top             =   1350
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   114
            Top             =   750
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFF0&
            Caption         =   "Option1"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   113
            Top             =   150
            Value           =   -1  'True
            Width           =   225
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   9
            Left            =   1320
            Picture         =   "SOLINT07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   5840
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   8
            Left            =   1320
            Picture         =   "SOLINT07.frx":1670
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   5220
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   7
            Left            =   1320
            Picture         =   "SOLINT07.frx":29D6
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   4620
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   6
            Left            =   1320
            Picture         =   "SOLINT07.frx":3D3C
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   4020
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   5
            Left            =   1320
            Picture         =   "SOLINT07.frx":50A2
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   3420
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   4
            Left            =   1320
            Picture         =   "SOLINT07.frx":6408
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   2840
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   3
            Left            =   1320
            Picture         =   "SOLINT07.frx":776E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   2230
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   2
            Left            =   1320
            Picture         =   "SOLINT07.frx":8AD4
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   1620
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   1
            Left            =   1320
            Picture         =   "SOLINT07.frx":9E3A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   1020
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            Height          =   250
            Index           =   0
            Left            =   1320
            Picture         =   "SOLINT07.frx":B1A0
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Búsqueda por Especificación"
            Top             =   420
            Width           =   435
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   9
            Left            =   8250
            Picture         =   "SOLINT07.frx":C506
            TabIndex        =   72
            Top             =   5520
            Width           =   175
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   8
            Left            =   8250
            Picture         =   "SOLINT07.frx":C810
            TabIndex        =   65
            Top             =   4920
            Width           =   175
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   9
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   71
            Text            =   " "
            Top             =   5520
            Width           =   450
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   8
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   64
            Text            =   " "
            Top             =   4920
            Width           =   450
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   7
            Left            =   8250
            Picture         =   "SOLINT07.frx":CB1A
            TabIndex        =   58
            Top             =   4320
            Width           =   175
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   6
            Left            =   8250
            Picture         =   "SOLINT07.frx":CE24
            TabIndex        =   51
            Top             =   3720
            Width           =   175
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   5
            Left            =   8250
            Picture         =   "SOLINT07.frx":D12E
            TabIndex        =   44
            Top             =   3120
            Width           =   175
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   4
            Left            =   8250
            Picture         =   "SOLINT07.frx":D438
            TabIndex        =   37
            Top             =   2520
            Width           =   175
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   7
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   57
            Text            =   " "
            Top             =   4320
            Width           =   450
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   6
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   50
            Text            =   " "
            Top             =   3720
            Width           =   450
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   5
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   43
            Text            =   " "
            Top             =   3120
            Width           =   450
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   4
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   36
            Text            =   " "
            Top             =   2520
            Width           =   450
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   3
            Left            =   8250
            Picture         =   "SOLINT07.frx":D742
            TabIndex        =   30
            Top             =   1920
            Width           =   175
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   2
            Left            =   8250
            Picture         =   "SOLINT07.frx":DA4C
            TabIndex        =   23
            Top             =   1320
            Width           =   175
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   3
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   29
            Text            =   " "
            Top             =   1920
            Width           =   450
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   2
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   22
            Text            =   " "
            Top             =   1320
            Width           =   450
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   1
            Left            =   8250
            Picture         =   "SOLINT07.frx":DD56
            TabIndex        =   16
            Top             =   720
            Width           =   175
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   1
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   15
            Text            =   " "
            Top             =   720
            Width           =   450
         End
         Begin VB.CommandButton Command19 
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
            Height          =   280
            Index           =   0
            Left            =   8250
            Picture         =   "SOLINT07.frx":E060
            TabIndex        =   9
            Top             =   120
            Width           =   175
         End
         Begin VB.TextBox Text4 
            Height          =   280
            Index           =   0
            Left            =   7820
            MaxLength       =   2
            TabIndex        =   8
            Text            =   " "
            Top             =   120
            Width           =   450
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
            Height          =   280
            Index           =   9
            Left            =   2640
            Picture         =   "SOLINT07.frx":E36A
            TabIndex        =   68
            Top             =   5520
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   9
            Left            =   345
            TabIndex        =   67
            Top             =   5520
            Width           =   2325
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
            Height          =   280
            Index           =   8
            Left            =   2640
            Picture         =   "SOLINT07.frx":E674
            TabIndex        =   61
            Top             =   4920
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   8
            Left            =   345
            TabIndex        =   60
            Top             =   4920
            Width           =   2325
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
            Height          =   280
            Index           =   7
            Left            =   2640
            Picture         =   "SOLINT07.frx":E97E
            TabIndex        =   54
            Top             =   4320
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   7
            Left            =   345
            TabIndex        =   53
            Top             =   4320
            Width           =   2325
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
            Height          =   280
            Index           =   6
            Left            =   2640
            Picture         =   "SOLINT07.frx":EC88
            TabIndex        =   47
            Top             =   3720
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   6
            Left            =   345
            TabIndex        =   46
            Top             =   3720
            Width           =   2325
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
            Height          =   280
            Index           =   5
            Left            =   2640
            Picture         =   "SOLINT07.frx":EF92
            TabIndex        =   40
            Top             =   3120
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   5
            Left            =   345
            TabIndex        =   39
            Top             =   3120
            Width           =   2325
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
            Height          =   280
            Index           =   4
            Left            =   2640
            Picture         =   "SOLINT07.frx":F29C
            TabIndex        =   33
            Top             =   2520
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   4
            Left            =   345
            TabIndex        =   32
            Top             =   2520
            Width           =   2325
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
            Height          =   280
            Index           =   3
            Left            =   2640
            Picture         =   "SOLINT07.frx":F5A6
            TabIndex        =   26
            Top             =   1920
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   3
            Left            =   345
            TabIndex        =   25
            Top             =   1920
            Width           =   2325
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
            Height          =   280
            Index           =   2
            Left            =   2640
            Picture         =   "SOLINT07.frx":F8B0
            TabIndex        =   19
            Top             =   1320
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   2
            Left            =   345
            TabIndex        =   18
            Top             =   1320
            Width           =   2325
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
            Height          =   280
            Index           =   1
            Left            =   2640
            Picture         =   "SOLINT07.frx":FBBA
            TabIndex        =   12
            Top             =   720
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   1
            Left            =   345
            TabIndex        =   11
            Top             =   720
            Width           =   2325
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
            Height          =   280
            Index           =   0
            Left            =   2640
            Picture         =   "SOLINT07.frx":FEC4
            TabIndex        =   5
            Top             =   120
            Width           =   175
         End
         Begin VB.TextBox Text5 
            Height          =   280
            Index           =   0
            Left            =   345
            TabIndex        =   4
            Top             =   120
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   0
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   120
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   0
            Left            =   8520
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   1
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   1
            Left            =   8520
            TabIndex        =   17
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   2
            Left            =   8520
            TabIndex        =   24
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   2
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   3
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   3
            Left            =   8520
            TabIndex        =   31
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   4
            Left            =   8520
            TabIndex        =   38
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   4
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   2520
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   5
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   5
            Left            =   8520
            TabIndex        =   45
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   6
            Left            =   8520
            TabIndex        =   52
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   6
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   3720
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   7
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            Top             =   4320
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   7
            Left            =   8520
            TabIndex        =   59
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   8
            Left            =   8520
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
            Top             =   4920
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   8
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Top             =   4920
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Index           =   9
            Left            =   2920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   5520
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   280
            Index           =   9
            Left            =   8520
            TabIndex        =   73
            Top             =   5520
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   9
            Left            =   8520
            TabIndex        =   152
            Top             =   5835
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   9
            Left            =   7800
            TabIndex        =   150
            ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
            Top             =   5880
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   8
            Left            =   8520
            TabIndex        =   149
            Top             =   5235
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   8
            Left            =   7800
            TabIndex        =   147
            ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
            Top             =   5280
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   7
            Left            =   8520
            TabIndex        =   146
            Top             =   4640
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   7
            Left            =   7800
            TabIndex        =   144
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   6
            Left            =   8520
            TabIndex        =   143
            Top             =   4030
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   6
            Left            =   7800
            TabIndex        =   141
            ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
            Top             =   4030
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   5
            Left            =   8520
            TabIndex        =   140
            Top             =   3435
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   5
            Left            =   7800
            TabIndex        =   138
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   4
            Left            =   8520
            TabIndex        =   137
            Top             =   2850
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   4
            Left            =   7800
            TabIndex        =   135
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   3
            Left            =   8520
            TabIndex        =   134
            Top             =   2220
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   3
            Left            =   7800
            TabIndex        =   132
            Top             =   2240
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   1
            Left            =   8520
            TabIndex        =   131
            Top             =   1035
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   1
            Left            =   7800
            TabIndex        =   129
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   2
            Left            =   8520
            TabIndex        =   128
            Top             =   1635
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   2
            Left            =   7800
            TabIndex        =   126
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
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
            Height          =   240
            Index           =   0
            Left            =   8520
            TabIndex        =   125
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "F. Ent.:"
            Height          =   195
            Index           =   0
            Left            =   7800
            TabIndex        =   123
            Top             =   440
            Width           =   735
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3480
         Left            =   9840
         Max             =   360
         TabIndex        =   92
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   10395
      ScaleHeight     =   6435
      ScaleWidth      =   1215
      TabIndex        =   81
      ToolTipText     =   "Visualizar Documentos Impresos"
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "SOLINT07.frx":101CE
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
         Top             =   3770
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Height          =   630
         Left            =   105
         Picture         =   "SOLINT07.frx":104D8
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Accesos Directos"
         Top             =   2110
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Height          =   640
         Left            =   105
         Picture         =   "SOLINT07.frx":107E2
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   1470
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Height          =   650
         Left            =   105
         Picture         =   "SOLINT07.frx":10AEC
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Listado de Sollicitudes por  Rango de Fechas"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "SOLINT07.frx":10DF6
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Anotaciones"
         Top             =   2940
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Height          =   630
         Left            =   100
         Picture         =   "SOLINT07.frx":11100
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Configuración de Funcionamiento General \Configuración de Formulario"
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         Height          =   650
         Left            =   100
         Picture         =   "SOLINT07.frx":11542
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         Height          =   650
         Left            =   100
         Picture         =   "SOLINT07.frx":1184C
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   100
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
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   82
         Top             =   5775
         Width           =   640
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
            Left            =   0
            TabIndex        =   83
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "SOLINT07.frx":11996
         Top             =   6030
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   150
      TabIndex        =   79
      Top             =   5880
      Width           =   6795
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Top             =   300
         Width           =   6360
      End
   End
   Begin VB.CommandButton Command20 
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
      Left            =   9990
      TabIndex        =   2
      Top             =   840
      Width           =   195
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7470
      TabIndex        =   168
      Top             =   405
      Width           =   1575
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   0
      TabIndex        =   167
      ToolTipText     =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   166
      ToolTipText     =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   0
      TabIndex        =   165
      ToolTipText     =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   0
      TabIndex        =   164
      ToolTipText     =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   0
      TabIndex        =   163
      ToolTipText     =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   1
      Left            =   0
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   2
      Left            =   0
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   3
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   4
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   5
      Left            =   0
      TabIndex        =   158
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   6
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   7
      Left            =   0
      TabIndex        =   156
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   8
      Left            =   0
      TabIndex        =   155
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   9
      Left            =   0
      TabIndex        =   154
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label17 
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   153
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   2520
      TabIndex        =   110
      ToolTipText     =   " "
      Top             =   7875
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   2520
      TabIndex        =   109
      ToolTipText     =   " "
      Top             =   7875
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   2520
      TabIndex        =   108
      ToolTipText     =   " "
      Top             =   7875
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2520
      TabIndex        =   107
      ToolTipText     =   " "
      Top             =   7875
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Imputacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   103
      Top             =   7710
      Width           =   975
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2520
      TabIndex        =   101
      ToolTipText     =   " "
      Top             =   7875
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   100
      ToolTipText     =   " "
      Top             =   7680
      Width           =   8955
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Solicitada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7440
      TabIndex        =   98
      Top             =   1365
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   9150
      TabIndex        =   97
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sector:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   96
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7995
      TabIndex        =   95
      ToolTipText     =   " "
      Top             =   1995
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "     Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8550
      TabIndex        =   90
      ToolTipText     =   " "
      Top             =   1995
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3180
      TabIndex        =   89
      Top             =   1995
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      TabIndex        =   88
      ToolTipText     =   " "
      Top             =   1995
      Width           =   855
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitud Interna de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   150
      TabIndex        =   86
      Top             =   210
      Width           =   6975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitante:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   78
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   1230
      TabIndex        =   77
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   9150
      TabIndex        =   76
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   9030
      TabIndex        =   75
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Emision:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7440
      TabIndex        =   74
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Solicitud:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Dobe-Click para Asignar Destino por Item"
      Top             =   405
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   150
      TabIndex        =   87
      Top             =   1920
      Width           =   10095
   End
End
Attribute VB_Name = "SOLINT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call SELECHO("SELFECHA") 'seleccion de fecha
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label11 = VALACT1$("SELFECHA")
   End If

End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    OPX% = ALTERNA%("Autorizaciones\Pedido de Cotizacion\Carga Cotizaciones Recibidas\Gestion de Compras\Anular SIM's")
    Select Case OPX%
        Case 1
           AUTORISIM.Show 1
        Case 2
           If DERACCE%("PEDICOT", "Pedidos de Cotizacion") > 1 Then FOPCMAQ07.Show 1
        Case 3
           If DERACCE%("CARCOTI", "Carga de Cotizacion") > 1 Then CARCOTI07.Show 1
        Case 4
           Call CONECRUN("ORCOMP07", "Ordenes de Compras")
        Case 5
            Call ANULASIM
    End Select
    Command10.Enabled = True
End Sub

Private Sub Command12_Click(Index As Integer)
   '
   Call SELECHO("SELFECHA") 'seleccion de fecha
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label19(Index) = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
    If TRIM$(Text6) = "" Then
       Call MENSERR(24, "Falta Sector Solicitante")
       GoTo 99
    End If
    '
    FEMI% = FECHANUM(Label3)
    FENT% = FECHANUM(Label11)
    '
    FF% = HOY(HO$) ' Validacion de que la fecha no sea menor a la del dia
    If FEMI% < FF% Then Call MENSERR(24, "La Fecha no Puede ser Menor a la de Hoy"): GoTo 99
    If FENT% < FF% Then Call MENSERR(24, "La Fecha de Entrega no Puede ser Menor a la de Hoy"): GoTo 99
    '
    If FEMI% > FENT% Then ' VALIDA QUE LA FECHA
       Call MENSERR(24, "La Fecha de Entrega debe ser Igual o Mayor a la Fecha de Emision")
       GoTo 99
    End If
    '
    VUELTA% = 0
    For I& = 0 To 9 'VALIDACION DE CODIGO
        If TRIM$(Text5(I&)) <> "" Or TRIM$(Text1(I&)) <> "" Or TRIM$(Text3(I&)) <> "" Then
           VUELTA% = VUELTA% + 1
           If TRIM$(Text5(I&)) <> "" Then  'SI CODIGO DIFERENTE DE VACIO
               If CODINT%(TRIM$(Text5(I&))) < 1 Then 'SI NO EXISTE EL CODIGO
                   Call MENSERR(24, "Código " & TRIM$(Text5(I&)) & " Inválido - Linea: " & I& + 1): GoTo 99
10                 If TRIM$(Text1(I&)) = "" Then 'SI  NO TIENE DESCRIPCION
                       Call MENSERR(24, "Falta Descripción Linea: " & I& + 1): GoTo 99
15                     If Text4(I&) = "" Or ECOARCH("UMEDIDA", TRIM$(Text4(I&))) = "" Then ' validacion de undiad de  medida
                           Call MENSERR(24, "Unidad de Medida no Válida Linea: " & I& + 1): GoTo 99
20                         If XVALO(Text3(I&)) < 0.05 Then 'SI TIENE CANTIDAD
                              Call MENSERR(24, "Falta Cantidad Linea: " & I& + 1): GoTo 99
25                              If FECHANUM(Label19(I&).Caption) < 1 Then ' VALIDA QUE TENGA FECHA CARGADAFECHA
                                    Label19(I&).Caption = FECHATEX$(FENT%)  ' SI NO TIENE FECHA CARGADA ASIGAN FECHA DE EMISION GENERAL
                                End If
                                If FECHANUM(Label19(I&).Caption) < FEMI% Then ' VALIDA QUE LA FECHA DE ENTREGA SEA MAYOR A LA FECHA DE EMISION
                                    Call MENSERR(24, "La Fecha de Entrega debe ser Igual o Mayor a la Fecha de Emision")
                                    GoTo 99
                                End If
                           Else
                              GoTo 25 ' VERIFICO FECHA DE EMISION Y DE ENTREGA
                           End If
                         Else
                           GoTo 20 ' VERIFICO SI TIENE  CANTIDAD
                       End If
                     Else
                       GoTo 15 ' VERIFICO LA UNIDAD DE MEDIDA
                   End If
                 Else
                   GoTo 10  ' VERIFICO SI TIENE DESCRIPCION
               End If
             Else
               GoTo 10
           End If
           '
           If ECOARCH$("ECUECON", TRIM$(Label15(I&))) = "" Then ' Validacion de Cuenta Contable
              Call MENSERR(24, "Imputacion no Válida - Linea: " & I& + 1)
              GoTo 99
           End If
        End If
        '
    Next I&
    '
    If VUELTA% < 1 Then ' VALIDO QUE SE HAYA INGRESADO AL MENOS UN DATO
       Call MENSERR(24, "No Se Han Ingresado Datos ")
       GoTo 99
    End If
    '
    Call ABRECOMPRAS
    XX& = CantRegistros("ENCASOLINT")
    If XX& < 1 Then
         PROSOLI& = 1
       Else
         PROSOLI& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", "NUME_SOLINT > 0")) 'incluir el proximo numero de solicitud
    End If
    If PROSOLI& > XVALO(Label28) Then   ' SI ES MAYOR EL NUMERO DE SOLICITUD DEL QUE ESTA EN PANTALLA
       Call COMUNI("FLEXOFT ha Re-Numerado la Solicitud de Materiales \como S/M Número '" + TRIM$(Str$(PROSOLI&)) + "'.\  \Verifique y Vuelva a Grabar.")
       Label28 = PROSOLI&
       GoTo 99
    End If
    '
    R& = USNBR%  'ASIGNACION DE LA CADENA DE AUTORIZANTES
    If R& >= 101 Then R& = R& Mod 100
    USER1% = Asc(Mid$(REGLEIDO$("AUSOLINT", R&), 32, 2)): US1& = USER1%
    USER2% = Asc(Mid$(REGLEIDO$("AUSOLINT", US1&), 32, 2)): US2& = USER2%
    USER3% = Asc(Mid$(REGLEIDO$("AUSOLINT", US2&), 32, 2)): US3& = USER3%
    USER4% = Asc(Mid$(REGLEIDO$("AUSOLINT", US3&), 32, 2)): US2& = USER2%
    '
    SQLX$ = "SELECT * FROM ENCASOLINT"  'GRABACION EN TABLA ENCASOLINT
    Dim ENCASTMP As ADODB.Recordset
    Set ENCASTMP = New ADODB.Recordset
    ENCASTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With ENCASTMP
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !NUME_SOLINT = PROSOLI&
       !FEMI_SOLINT = Label3
       !Usua_Gen = USNBR% Mod 100
       !Obse_Solint = Text2
       !AUTO_1 = USER1% Mod 100
       !AUTO_2 = USER2% Mod 100
       !AUTO_3 = USER3% Mod 100
       !AUTO_4 = USER4% Mod 100
       !Sect_Solint = Left$(Text6, 12)
       !HayNews = 1
       DESCRIP$ = Label3 + " - " + Left$(Text6, 12) + " - " + Text1(0).TEXT
       !DESCRI_SOLINT = Left$(DESCRIP$, 64)
       .Update
    End With
    ENCASTMP.Close
    Set ENCASTMP = Nothing
    '
    J% = 0
    SQLX$ = "SELECT * FROM DETASOLINT"  'GRABACION EN TABLA DETASOLINT
    Dim DETASTMP As ADODB.Recordset
    Set DETASTMP = New ADODB.Recordset
    DETASTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With DETASTMP
      For U& = 1 To 10
        If Text1(U& - 1) <> "" Then ' SI TIENE DESCRIPCION SE GRABA
          J% = J% + 1
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUME_SOLINT = PROSOLI&
          !NUORD_ITEM = J%
              CI% = CODINT(Text5(U& - 1))
          !COD_INTE = CI%
          !CANT_SOLINT = XVALO(Text3(U& - 1))
          !DESCRITEM = Text1(U& - 1)
          !Uni_Mas = UNIMED$(CI%)
          !UNI_SOL = Text4(U& - 1)
          !Fentre_Sol = Label19(U& - 1)
          !DESTI_SOLINT = Left$(Label17(U& - 1), 255)
          !Iconta_Solint = Left$(Label15(U& - 1), 32)
          !Status = 0
          .Update
        End If
      Next U&
    End With
    DETASTMP.Close
    Set DETASTMP = Nothing
    '
    Call PRINTSOLI
    '
    TTXX$ = "Generación de SIM" & CRLF$
    TTXX$ = TTXX$ + "Sector: " & Text6.TEXT & CRLF$
    TTXX$ = TTXX$ + "Solicitante: " & Label4.Caption & CRLF$
    For I& = 0 To 9
        If TRIM$(Text5(I&)) = "" And TRIM$(Text1(I&)) = "" Then GoTo 30
        TTXX$ = TTXX$ + CRLF$
        TTXX$ = TTXX$ & "Item Numero: " & (I& + 1) & CRLF$
        TTYY$ = TRIM$(Text5(I&))
        If TTYY$ <> "" Then TTYY$ = TTYY$ + " - "
        TTYY$ = TTYY$ & TRIM$(Text1(I&)) & CRLF$
        TTXX$ = TTXX$ + TTYY$
        TTXX$ = TTXX$ + "Cantidad: " & TRIM$(Text3(I&)) & " " & TRIM$(Text4(I&)) & CRLF$
30  Next I&
    TTXX$ = TTXX$ + CRLF$
    TTXX$ = TTXX$ + "Observaciones: " & Text2.TEXT
    '
    Call ANOTASIM07.ANOTASIM(XVALO(Label28.Caption), TTXX$)
    '
    Call BLANFORMU
    '
    PROSOLI& = XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", "NUME_SOLINT > 0")) 'incluir el proximo numero de solicitud
    Label28 = PROSOLI& + 1 'SIGUIENTE NUMERO DE SOLICITUD
    '
    VScroll1.Value = 0
    '
    Call GENINSOL
    Call LOCINSOL(USNBR%)
    '
99  Command14.Enabled = True
End Sub

Private Sub Command19_Click(Index As Integer)
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4(Index).TEXT = VALONUE$
   End If

End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Command20_Click()
   Call SELECHO("SELFECHA") 'seleccion de fecha
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label3 = VALACT1$("SELFECHA")
   End If

End Sub

Private Sub Command3_Click(Index As Integer)
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda de Especificaciones", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     '
     Call SETULTREG("!ECOMAST", 0)
     Call APERBASDAT("MASTER")
     CONDI$ = "Especificacion LIKE '%" & ARGU$ + "%'"
     CONDI$ = CONDI$ + " AND STATUS > 0 "
     CAMPOS$ = "CODIGO/A16;DESCRIPCION/A46"
     Call CARGALISEL("!ECOMAST", "MASTER", CAMPOS$, CONDI$)
'     SQLX$ = "Select * from MASTER "
'     SQLX$ = SQLX$ + "WHERE Especificacion LIKE '%" & ARGU$ + "%'"
'     SQLX$ = SQLX$ + " AND STATUS > 0 "
'     Dim MASTERTMP As ADODB.Recordset
'     Set MASTERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'     With MASTERTMP
'       Do While Not .EOF
'         JJ& = JJ& + 1
'         CI% = XVALO(!CODINT)
'         CODD$ = CODEXT$(CI%)
'         DESC$ = DESCRIT$(CI%)
'         Call REPLA(TX$, CODD$, 1, 16)
'         Call REPLA(TX$, DESC$, 17, 46)
'         Call GRAREG("!ECOMAST", TX$, JJ&)
'         .MoveNext
'       Loop
'     End With
'     MASTERTMP.Close
'     Set MASTERTMP = Nothing
'     Screen.MousePointer = 1
'     Call SETULTREG("!ECOMAST", JJ&)
'     Call FRESHECHO("!ECOMAST")
   '
    If ULTREG&("!ECOMAST") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
       Call SELECHO("!ECOMAST")
       Text5(Index) = VALACT1$("!ECOMAST")
     Else
       Call MENSERR(24, "No Se Encontraron Registros que Cumplan con la Consulta SQL")
       GoTo 10
     End If
  End If
End Sub

Sub Command4_Click()
    '
    Command4.Enabled = False
    '
10  If ULTREG&("!INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
     Call SELECHO("!INSOLINT/Indice General de Solicitud de Materiales")
    Else
     GoTo 99
    End If
    '
    NUDOCU$ = VALACT1$("!INSOLINT") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    TIDOCUM$ = "Solicitud de Materiales"
    If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
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
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
    GoTo 10
       '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
   GoTo 10

99  Command4.Enabled = True

End Sub

Private Sub Command5_Click()
    Command5.Enabled = False 'OPCIONES DE CONFIGURACION
    '
    RTA1% = ALTERNA%("Configuración de Formulario \Tabla de Autorizantes \Tabla de Imputación de Solicitudes \Configuración General\Generador de Reportes")
    If RTA1% < 1 Or RTA1% > 5 Then GoTo 99
    If RTA1% = 1 Then ' CONFIGURACION DE FORMULARIO
        If DERACCE%("FORMSIM", "Configuracion de Formulario de SIM") > 1 Then FOSOLMAT.Show 1
    ElseIf RTA1% = 2 Then  ' TABLA DE AUTORIZANTES
     If DERACCE%("AUSOLINT", "Acceso a la Tabla de Autorizantes") > 1 Then
        Call TRALOCAL("AUSOLINT", "")
        '
        N1% = FILEOPEN%("USER.DAT", 9, 64)
        Dim F1 As String * 64
        '
        Y& = 0
        For I& = 1 To LOF(N1%) / 64
           '
           X$ = REGLEIDO$("!AUSOLINT", I&)
           If I& > ULTREG&("!AUSOLINT") Then X$ = REGBLAN$("AUSOLINT")
           I1% = I&
           Call REPLA(X$, Chr$(I1%), 1, 1)
           NOM$ = USERNAM(I1%)
           Call REPLA(X$, UCase$(NOM$), 2, 30)
           If I& > ULTREG&("!AUSOLINT") Then
             Call REPLA(X$, Chr$(0), 32, 1)
             Call REPLA(X$, MKI$(0), 33, 2)
           End If
           Call GRAREG("!AUSOLINT", X$, I&)
        Next I&
        '
        Close #N1%: N1% = 0
        Call CIERRARCH("!AUSOLINT")
        '
        VTB% = VENTABU%("AUSOLINT")
        If VTB% < 0 Then GoTo 99

        Call TRACENTRAL("AUSOLINT", "")
      End If
    ElseIf RTA1% = 3 Then  'Tabla de Imputación de Solicitudes
      '
      If DERACCE%("TABAUTO", "Tabla de Imputacion de Solicitudes") > 1 Then
        OPX% = COMALTER%("Forma de Actualización:\\Directa\Importación de Planilla Excel")
        If OPX% = 1 Then
             Call TRALOCAL("IMPUSOL", "")
             VTB% = VENTABU%("IMPUSOL")
             If VTB% < 0 Then GoTo 99
             Call TRACENTRAL("IMPUSOL", "")
           ElseIf OPX% = 2 Then
             IMPOIMPU07.Show 1
        End If
      End If
      '
    ElseIf RTA1% = 4 Then  ' CONFIGURACION GENERAL
      Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
      Call FINAL("")
    ElseIf RTA1% = 5 Then  ' CONFIGURACION GENERAL
      Call CONECRUN("GENREP07", "Generador de Reportes")
    End If
    '
99  Command5.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = True
    If DERACCE%("LISTASIM", "Listado de Solicitudes de Materiales") > 1 Then
      RTA1% = ALTERNA%("Solicitudes Por Número\Solicitud por Código\Autorizaciones\Destino Contable")
      If RTA1% < 1 Or RTA1% > 4 Then GoTo 99 ' LISTADOS EN EL GENRADOR DE REPORTE
      If RTA1% = 1 Then Call FINAL("?SIMPORNUMERO")
      If RTA1% = 2 Then Call FINAL("?SIMPORCODIGO")
      If RTA1% = 3 Then Call FINAL("?LAUTORISIM")
      If RTA1% = 4 Then Call FINAL("?SIMCONTABLE")
    End If
99  Command6.Enabled = True
 
End Sub


Private Sub Command7_Click()
   Command7.Enabled = False
   ANOTASIM07.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("AMASTO07", "Maestro de Items de Stock")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click(Index As Integer)
   Call SELECHO("!ECOMASIM")
   VALONUE$ = VALACT1$("!ECOMASIM")
   If VALONUE$ <> "" Then
      Text5(Index).TEXT = VALONUE$
   End If

End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If DERACCE%("SOLINT", "Solicitud Interna de Materiales") < 2 Then
      Call FINAL("") ' si no tiene derechos se  cierra el formulario
    End If
    '
    Call ABRECONEXION
    Label4.Caption = USERNAM$(USNBR%) 'asiga el nombre del usuario
    '
    Call LOCINSOL(USNBR%)
    '
    Dim CLACLA$(2048), VECES%(2048)
    CANCLA% = 0
    '
10  TreeView1.Nodes.Add , , "C00", "IMPUTACION" ' CREO EL NODO
    For I& = 1 To ULTREG("IMPUSOL")
       X$ = REGLEIDO("IMPUSOL", I&)
       CLAIMPU$ = TRIM$(Left$(X$, 12))
       If Len(CLAIMPU$) >= 2 Then
         PADRE$ = Left$(CLAIMPU$, Len(CLAIMPU$) - 2)
         PP$ = "C" & PADRE$
         If PADRE$ = "" Then PP$ = "C00"
         NN$ = "C" & CLAIMPU$
         CCONT$ = TRIM$(Mid$(X$, 77, 12))
         If CCONT$ <> "" Then NN$ = CCONT$
         '
         For KKII% = 1 To CANCLA%
            If CLACLA$(KKII%) = NN$ Then
               VECES%(KKII%) = VECES%(KKII%) + 1
               NN$ = AJUSTI$(NN$, 16) + "/" + TRIM$(Str$(VECES%(KKII%)))
               GoTo 17
            End If
         Next KKII%
         CANCLA% = CANCLA% + 1
         CLACLA$(CANCLA%) = NN$
         '
17       REFER$ = TRIM$(Mid$(X$, 13, 64))
         On Error Resume Next
         ' Esto se agrega porque la clave que se pasa al treeview debe ser alfanumerica
         ' en caso contrario no la agrega al arbol
         If CCONT$ <> "" Then NN$ = AJUSTI(NN$, 20) + "/"
         TreeView1.Nodes.Add PP$, tvwChild, NN$, REFER$
         On Error GoTo 0
       End If
    Next I&
    '
    FF% = HOY(HO$)
    Label3 = HO$
    Label11 = HO$
    '
    XX& = CantRegistros("ENCASOLINT")
    If XX& < 1 Then
         PROSOLI& = 1
       Else
         PROSOLI& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", "NUME_SOLINT > 0")) 'incluir el proximo numero de solicitud
    End If
    Label28 = PROSOLI&
    '
    Call GENECOMASIM
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call FINAL("")
End Sub



Private Sub Option1_Click(Index As Integer)
    ITACTIVO% = Index
    Label13 = Label17(ITACTIVO%)
End Sub

Private Sub Text3_LostFocus(Index As Integer)
   Text3(Index).TEXT = FORMATNUM$(Text3(Index).TEXT, "F10.1") 'Formatea el valor de la cantidad con 1 decimal
End Sub


Private Sub Text4_DblClick(Index As Integer)
   Call Command19_Click(Index)
End Sub

Private Sub Text5_Change(Index As Integer)

  CODD$ = TRIM$(Text5(Index))
  CI% = CODINT%(CODD$) ' asigna codigo interno
  Text1(Index).TEXT = DESCRIT$(CI%) 'busca descripcion si el codigo existe, si no existe asiga Material No Codificado
  Text4(Index) = UNIMED(CI%)
End Sub

Private Sub Text5_DblClick(Index As Integer)
   Call Command9_Click(Index)
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    ITACTIVO% = (-1)
    For KKI% = 0 To 9
       If Option1(KKI%).Value = True Then
          ITACTIVO% = KKI%
          GoTo 10
       End If
    Next KKI%
    Call MENSERR(24, "Seleccione Registro Activo.")
    Exit Sub
    '
10  Dim n As Integer
    Dim strParents As String    ' Variable para información.
    n = Node.Index  ' Índice del nodo en el que se ha hecho clic.
    NUCO$ = TRIM$(Left$(Node.Key, 16)) ' CLAVE
    CUCO% = CUECOINT(NUCO$)
    If CUCO% > 0 Then
        Label15(ITACTIVO%) = NUCO$ ' PRESENTA LA CUENTA CONTABLE EN EL LABEL
    Else
        Label15(ITACTIVO%) = "" ' SI NO LIMPIA
        Label13 = ""
    End If
    strParents = Node.TEXT & vbLf
    While n <> Node.Root.Index
        strParents = TreeView1.Nodes(n).Parent.TEXT + " - " + strParents
        ' Establece a n el índice del siguiente nodo primario.
        n = TreeView1.Nodes(n).Parent.Index
    Wend
    '
    If Left$(strParents, 10) = "IMPUTACION" Then strParents = Mid$(strParents, 14)
    Label13 = strParents ' RUTA
    Label17(ITACTIVO%) = strParents
    '
End Sub

Private Sub VScroll1_Change()
   If VScroll1.Value < 280 Then
     Picture5.Top = (-10) * VScroll1.Value
   End If
End Sub

Private Sub VScroll1_Scroll()
   If VScroll1.Value < 280 Then
    Picture5.Top = (-10) * VScroll1.Value
   End If
End Sub

Sub BLANFORMU()
  For J& = 0 To 9
    Text5(J&) = ""
    Text3(J&) = ""
    Text1(J&) = ""
    Text4(J&) = ""
    Label19(J&) = ""
  Next J&
  Text2 = ""
  '
  FF% = HOY(HO$) ' presenta la fecha del dia por defecto
  Label3 = HO$
  Label11 = HO$
  Text6 = "" ' limpia el text de  sector
  Label13 = ""
  Option1(0).Value = True
  '
End Sub

Sub PRINTSOLI()
    '
    CANCELPRINT% = 0 ' VERIFICA SI TIENE SUPRIMIDA LA VALIDACION
    COPRISIM.Show 1
    If COPRISIM.Check2.Value = 1 Then CANCELPRINT = 1
    '
    TIPODOC$ = "Solicitud de Materiales" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Label28)
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = Label3 'FECHA
    CODPARAM$(2) = "NUME-COM": DOCPARAM$(2) = TRIM$(Label28) 'NUMERO DE SOLICITUD
    CODPARAM$(3) = "OBSERVA":  DOCPARAM$(3) = Text2.TEXT   'Observaciones
    CODPARAM$(4) = "REF-MAT1":  DOCPARAM$(4) = Left$(TRIM$(Text6), 12) 'SECTOR
    'CODPARAM$(5) = "FECH-VEN":  DOCPARAM$(5) = Label11   'FECHA SOLICITADA
    '
    NUMAUT1% = VALOBADA("ENCASOLINT", "Auto_1", "Nume_Solint = " & TRIM$(Label28) & "") 'AUTORIZANTE 1
    NUMAUT2% = VALOBADA("ENCASOLINT", "Auto_2", "Nume_Solint = " & TRIM$(Label28) & "") 'AUTORIZANTE 2
    NUMAUT3% = VALOBADA("ENCASOLINT", "Auto_3", "Nume_Solint = " & TRIM$(Label28) & "") 'AUTORIZANTE 3
    '
    CODPARAM$(5) = "MEDIDA-1":  DOCPARAM$(6) = USERNAM$(NUMAUT1%)
    CODPARAM$(6) = "MEDIDA-2":  DOCPARAM$(7) = USERNAM$(NUMAUT2%)
    CODPARAM$(7) = "MEDIDA-3":  DOCPARAM$(8) = USERNAM$(NUMAUT3%)
    
    CAPARDOC% = 7
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "ORD-ITEM"
    CODTABU1$(6) = "REF-MAT2"
    CODTABU1$(7) = "FECH-VEN"
    CACOLTAB1% = 7
    '
    CAITAB1% = 0
    '
    FORIPRE$ = CONTROL$("", "FOSOLMAT")
    FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
    ATELIB% = XVALO(Mid$(FORTELIB$, 2))
    If ATELIB% < 16 Then ATELIB% = 16
    CAITAB1% = 0
    For U& = 1 To 10
       If TRIM$(Text1(U& - 1)) <> "" Then
        CAITAB1% = CAITAB1% + 1
        J& = J& + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Text5(U& - 1))  ' CODIGO DE MATERIAL
        TETABU1$(3, CAITAB1%) = TRIM$(Text3(U& - 1))   'CANTIDAD
        TETABU1$(4, CAITAB1%) = TRIM$(Text4(U& - 1))   'UNIDAD DE MEDIDA
        TETABU1$(5, CAITAB1%) = J&   'NUMERO DE ORDEN
        TETABU1$(6, CAITAB1%) = Label15(U& - 1) 'IMPUTACION CONTABLE
        TETABU1$(7, CAITAB1%) = Label19(U& - 1) 'FECHA SOLICITADA POR ITEM
        '
        TELIBRE$ = RTrim(Text1(U& - 1))
        Call FRATEXTO(TELIBRE$, ATELIB%)
        TETABU1$(2, CAITAB1%) = RETEX$(1)
        For KKU1% = 2 To CARETEX%
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        '
       End If
    Next U&
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOSOLMAT"))
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GENINSOL() 'CARGA ARCHIVO CENTRAL DE SOLICITUDES INTERNAS
    '
    SQLX$ = "SELECT * FROM ENCASOLINT"
    SQLX$ = SQLX$ + " ORDER BY Nume_Solint "
    Dim ENCATMP As ADODB.Recordset
    Set ENCATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ENCATMP
        JJ& = 0
       Do While Not .EOF
            XXY$ = REGBLAN$("INSOLINT")
            REFER$ = SAFETEXT(!DESCRI_SOLINT)
            Call REPLA(XXY$, MKS(!NUME_SOLINT), 1, 4)
            If XVALO(!Status) = 9 Then
                REFER$ = "(Anul) " + REFER$
            End If
            Call REPLA(XXY$, REFER$, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("INSOLINT", XXY$, JJ&)
            .MoveNext
       Loop
    End With
    ENCATMP.Close
    Set ENCATMP = Nothing
    Call SETULTREG("INSOLINT", JJ&)
    Call CIERRARCH("INSOLINT")
    Call FRESHECHO("INSOLINT")
    '
End Sub

Sub LOCINSOL(USER%) 'CARGA ARCHIVO LOCAL DE SOLICITUDES INTERNAS
    '
    USS% = USER% Mod 100
    SQLX$ = "SELECT * FROM ENCASOLINT"
    SQLX$ = SQLX$ + " where (Status <> 9 OR STATUS IS NULL) " 'para no tener en cuenta las sim anuladas
    SQLX$ = SQLX$ + " and (USUA_GEN = " & USS%
    SQLX$ = SQLX$ + " OR AUTO_1 = " & USS%
    SQLX$ = SQLX$ + " OR AUTO_2 = " & USS%
    SQLX$ = SQLX$ + " OR AUTO_3 = " & USS%
    SQLX$ = SQLX$ + " OR AUTO_4 = " & USS%
    SQLX$ = SQLX$ + ") ORDER BY Nume_Solint "
    Dim ENCATMP As ADODB.Recordset
    '
    JJ& = 0
    On Error Resume Next
    Set ENCATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Err Then
       On Error GoTo 0
       GoTo 20
    End If
    On Error GoTo 0
    With ENCATMP
        JJ& = 0
       Do While Not .EOF
            XXY$ = REGBLAN$("INSOLINT")
            REFER$ = SAFETEXT(!DESCRI_SOLINT)
            Call REPLA(XXY$, MKS(!NUME_SOLINT), 1, 4)
            Call REPLA(XXY$, REFER$, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INSOLINT", XXY$, JJ&)
            .MoveNext
       Loop
    End With
    ENCATMP.Close
    Set ENCATMP = Nothing
20  Call SETULTREG("!INSOLINT", JJ&)
    Call CIERRARCH("!INSOLINT")
    Call FRESHECHO("!INSOLINT")
    '
End Sub

Sub GENECOMASIM() ' GENERA LISTA DE SELECCION DE ARTICULOS POR TIPO DE MATERIAL

     Call COPYSTRU("ECOMAST", "ECOMASIM")
     '
     Call SETULTREG("!ECOMASIM", 0)
     Call APERBASDAT("MASTER")
     TIMAT$ = CONTROL(IDENTER$, "FITIPOMA") 'ASIGNA EL TIPO DE MATERIAL CONFIGURADO EN LA ESTACION DE TRABAJO
     CONDI$ = "STATUS > 0 "
     If TIMAT$ <> "" Then
        SQLX$ = SQLX$ + "AND TIMATE = " & TIMAT$
     End If
     CONDI$ = CONDI$ + " order by CODIGO"
     CAMPOS$ = "CODIGO/A16;DESCRIPCION/A46"
     Call CARGALISEL("!ECOMASIM", "MASTER", CAMPOS$, CONDI$)
     
'     SQLX$ = "Select * from MASTER "
'     SQLX$ = SQLX$ + "WHERE STATUS > 0 "
'     If TIMAT$ <> "" Then
'        SQLX$ = SQLX$ + "AND TIMATE = " & TIMAT$
'     End If
'     SQLX$ = SQLX$ + " order by CODIGO"
'     Dim MASTERTMP As ADODB.Recordset
'     Set MASTERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'     JJ& = 0
'     With MASTERTMP
'       Do While Not .EOF
'         JJ& = JJ& + 1
'         CI% = XVALO(!CODINT)
'         CODD$ = CODEXT$(CI%)
'         DESC$ = DESCRIT$(CI%)
'         Call REPLA(TX$, CODD$, 1, 16)
'         Call REPLA(TX$, DESC$, 17, 46)
'         Call GRAREG("!ECOMASIM", TX$, JJ&)
'         .MoveNext
'       Loop
'     End With
'     MASTERTMP.Close
'     Set MASTERTMP = Nothing
'     Screen.MousePointer = 1
'     Call SETULTREG("!ECOMASIM", JJ&)
'     Call FRESHECHO("!ECOMASIM")

End Sub

Sub ANULASIM()

    
    If ULTREG&("INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
         Call SELECHO("INSOLINT/Indice SIM's")
    Else
         Call MENSERR(24, "No Existen SIM's que Puedan Anularse")
         Exit Sub
    End If
    '
    SIMNUM& = XVALO(VALACT1$("INSOLINT"))
    If SIMNUM& < 1 Then
        Exit Sub
    End If
    '
    OPX% = COMALTER%("¿Realmente Desea Anular la SIM " & SIMNUM& & "?\\No, Conservar\Si, Anular")
    If OPX% = 2 Then
        Call ABRECOMPRAS
        Call ACTUREGISTRO("ENCASOLINT", "Status", "NUME_SOLINT = " & SIMNUM&, 9, REG&)
        Call ACTUREGISTRO("DETASOLINT", "Status", "NUME_SOLINT = " & SIMNUM&, 9, REG&)
        Call GENINSOL
        Call MENSERR(24, "SIM " & SIMNUM& & " Anulada")
        TTXX$ = "Anulada por " & USERNAME$ & " - " & TIMECENT$()
        Call ANOTASIM07.ANOTASIM(SIMNUM&, TTXX$)
    End If

End Sub




