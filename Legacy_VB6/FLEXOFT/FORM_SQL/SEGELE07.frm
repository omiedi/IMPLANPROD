VERSION 5.00
Begin VB.Form SEGELE07 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Ensayo de Fuga a Tierra"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13200
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
   ScaleHeight     =   8655
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text19 
      Height          =   405
      Index           =   0
      Left            =   7755
      TabIndex        =   43
      Top             =   10440
      Width           =   795
   End
   Begin VB.TextBox text19 
      Height          =   405
      Index           =   1
      Left            =   7995
      TabIndex        =   42
      Top             =   10845
      Width           =   795
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   800
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   2685
      Width           =   3855
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   460
         Width           =   2760
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realizó:"
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
         Left            =   50
         TabIndex        =   36
         Top             =   480
         Width           =   800
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
         Height          =   240
         Left            =   80
         TabIndex        =   35
         Top             =   195
         Width           =   800
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   34
         Top             =   150
         Width           =   2760
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   12015
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4200
         TabIndex        =   110
         Top             =   240
         Width           =   550
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
         Height          =   285
         Left            =   4755
         TabIndex        =   101
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label98 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4935
         TabIndex        =   102
         ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analizador de Corriente de Fuga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   9240
         TabIndex        =   31
         Top             =   360
         Width           =   2660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   9240
         TabIndex        =   30
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "INSTRUMENTOS UTILIZADOS EN ENSAYO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   80
         TabIndex        =   29
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   800
      Index           =   1
      Left            =   4080
      TabIndex        =   26
      Top             =   2685
      Width           =   3975
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IT -  42"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   47
         Top             =   510
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Corriente de Fuga del Aparato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   30
         TabIndex        =   27
         Top             =   240
         Width           =   3900
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5520
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   12015
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1080
         Picture         =   "SEGELE07.frx":0000
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1080
         Picture         =   "SEGELE07.frx":030A
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label92 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "30/05/16"
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
         Index           =   2
         Left            =   10440
         TabIndex        =   25
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label92 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R117"
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
         Index           =   1
         Left            =   11160
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ensayo de Seguridad Eléctrica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   0
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.TextBox Text9 
      Height          =   700
      Left            =   8100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   4050
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
      Height          =   8715
      Left            =   12240
      ScaleHeight     =   8655
      ScaleWidth      =   870
      TabIndex        =   7
      Top             =   360
      Width           =   930
      Begin VB.CommandButton Command7 
         Height          =   650
         Left            =   105
         Picture         =   "SEGELE07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Tabla de Instrumentos"
         Top             =   3045
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   105
         Picture         =   "SEGELE07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Visualiza e Imprime Informe de  Fuga a Tierra"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Height          =   650
         Left            =   105
         Picture         =   "SEGELE07.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   650
         Left            =   120
         Picture         =   "SEGELE07.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Graba e Imprime Informe"
         Top             =   6600
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
         TabIndex        =   8
         Top             =   7560
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
            TabIndex        =   9
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Height          =   630
         Left            =   105
         Picture         =   "SEGELE07.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "SEGELE07.frx":181E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SEGELE07.frx":1968
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7845
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   12015
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   10665
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   240
         Width           =   705
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   3675
         Left            =   560
         TabIndex        =   48
         Top             =   600
         Width           =   10815
         Begin VB.TextBox text41 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   3120
            TabIndex        =   109
            Text            =   " "
            Top             =   2790
            Width           =   1050
         End
         Begin VB.TextBox text41 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   3120
            TabIndex        =   108
            Text            =   " "
            Top             =   2490
            Width           =   1050
         End
         Begin VB.TextBox text41 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   3120
            TabIndex        =   107
            Text            =   " "
            Top             =   2205
            Width           =   1050
         End
         Begin VB.TextBox text41 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   3120
            TabIndex        =   105
            Text            =   " "
            Top             =   1905
            Width           =   1050
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   6300
            TabIndex        =   56
            Text            =   " "
            Top             =   1905
            Width           =   1500
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   6300
            TabIndex        =   55
            Text            =   " "
            Top             =   2205
            Width           =   1500
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   6300
            TabIndex        =   54
            Text            =   " "
            Top             =   2490
            Width           =   1500
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   6300
            TabIndex        =   53
            Text            =   " "
            Top             =   2790
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   0
            Left            =   9300
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   " "
            Top             =   1905
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   1
            Left            =   9300
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   " "
            Top             =   2205
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   2
            Left            =   9300
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   " "
            Top             =   2490
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   3
            Left            =   9300
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   " "
            Top             =   2790
            Width           =   1500
         End
         Begin VB.Label Labe141 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "---"
            Height          =   300
            Index           =   1
            Left            =   3105
            TabIndex        =   106
            Top             =   2205
            Width           =   1050
         End
         Begin VB.Label Label21 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "(N) Normal   (R) Reversa"
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
            Left            =   960
            TabIndex        =   98
            Top             =   3405
            Width           =   9855
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "(CN) Condición Normal   (TA) Tierra abierta"
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
            Left            =   960
            TabIndex        =   97
            Top             =   3120
            Width           =   9855
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "(2)"
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
            Left            =   0
            TabIndex        =   96
            Top             =   3390
            Width           =   975
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "(1)"
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
            Left            =   0
            TabIndex        =   95
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label label138 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TA"
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   94
            Top             =   2205
            Width           =   1065
         End
         Begin VB.Label label138 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CN"
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   93
            Top             =   1905
            Width           =   1065
         End
         Begin VB.Label Labe138 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TA"
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   92
            Top             =   2490
            Width           =   1065
         End
         Begin VB.Label label138 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CN"
            Height          =   300
            Index           =   3
            Left            =   960
            TabIndex        =   91
            Top             =   2790
            Width           =   1065
         End
         Begin VB.Label Labe143 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   1
            Left            =   5235
            TabIndex        =   90
            Top             =   2205
            Width           =   1050
         End
         Begin VB.Label Labe143 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            Height          =   300
            Index           =   0
            Left            =   5235
            TabIndex        =   89
            Top             =   1905
            Width           =   1050
         End
         Begin VB.Label Labe143 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   2
            Left            =   5235
            TabIndex        =   88
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Labe143 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100"
            Height          =   300
            Index           =   3
            Left            =   5235
            TabIndex        =   87
            Top             =   2790
            Width           =   1050
         End
         Begin VB.Label Labe144 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "50"
            Height          =   300
            Index           =   1
            Left            =   7800
            TabIndex        =   86
            Top             =   2205
            Width           =   1500
         End
         Begin VB.Label Labe144 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10"
            Height          =   300
            Index           =   0
            Left            =   7800
            TabIndex        =   85
            Top             =   1905
            Width           =   1500
         End
         Begin VB.Label Labe144 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "50"
            Height          =   300
            Index           =   2
            Left            =   7800
            TabIndex        =   84
            Top             =   2490
            Width           =   1500
         End
         Begin VB.Label Labe144 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10"
            Height          =   300
            Index           =   3
            Left            =   7800
            TabIndex        =   83
            Top             =   2790
            Width           =   1500
         End
         Begin VB.Label Label137 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            Height          =   300
            Index           =   3
            Left            =   0
            TabIndex        =   82
            Top             =   2790
            Width           =   945
         End
         Begin VB.Label Label137 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   81
            Top             =   2490
            Width           =   945
         End
         Begin VB.Label Labe137 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   80
            Top             =   2205
            Width           =   945
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tensión de Alimentación y Frecuencia: "
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
            Left            =   0
            TabIndex        =   79
            Top             =   360
            Width           =   10770
         End
         Begin VB.Label Labe141 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "---"
            Height          =   300
            Index           =   3
            Left            =   3105
            TabIndex        =   78
            Top             =   2790
            Width           =   1050
         End
         Begin VB.Label Labe142 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   3
            Left            =   4170
            TabIndex        =   77
            Top             =   2790
            Width           =   1050
         End
         Begin VB.Label Labe142 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   2
            Left            =   4170
            TabIndex        =   76
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Labe142 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   0
            Left            =   4170
            TabIndex        =   75
            Top             =   1905
            Width           =   1050
         End
         Begin VB.Label Labe141 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "---"
            Height          =   300
            Index           =   0
            Left            =   3105
            TabIndex        =   74
            Top             =   1905
            Width           =   1050
         End
         Begin VB.Label Labe141 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "---"
            Height          =   300
            Index           =   2
            Left            =   3105
            TabIndex        =   73
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Labe140 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            Height          =   300
            Index           =   2
            Left            =   2040
            TabIndex        =   72
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Labe142 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "500"
            Height          =   300
            Index           =   1
            Left            =   4170
            TabIndex        =   71
            Top             =   2205
            Width           =   1050
         End
         Begin VB.Label Labe140 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R"
            Height          =   300
            Index           =   3
            Left            =   2040
            TabIndex        =   70
            Top             =   2790
            Width           =   1050
         End
         Begin VB.Label Labe140 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "N"
            Height          =   300
            Index           =   0
            Left            =   2040
            TabIndex        =   69
            Top             =   1905
            Width           =   1050
         End
         Begin VB.Label Labe140 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "N"
            Height          =   300
            Index           =   1
            Left            =   2040
            TabIndex        =   68
            Top             =   2205
            Width           =   1050
         End
         Begin VB.Label Label137 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   67
            Top             =   1905
            Width           =   945
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "        Veredicto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   9300
            TabIndex        =   66
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corriente de Fuga Máxima Permitida (uA-Dc)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   7800
            TabIndex        =   65
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corriente de Fuga Máxima medida (uA-Dc)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   6300
            TabIndex        =   64
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corriente de Fuga Máxima Permitida (uA-Ac) Clase ||"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   5220
            TabIndex        =   63
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corriente de Fuga Máxima Permitida (uA-Ac)  Clase |"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   4170
            TabIndex        =   62
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corriente de Fuga Máxima medida (uA-Ac)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   3120
            TabIndex        =   61
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Polaridad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   2040
            TabIndex        =   60
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "             Fila N."
            Height          =   1005
            Left            =   0
            TabIndex        =   59
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Condición de Ensayo (1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   960
            TabIndex        =   58
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ensayo: Corriente de Fuga a Tierra"
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
            Left            =   0
            TabIndex        =   57
            Top             =   105
            Width           =   10815
         End
      End
      Begin VB.Label Label10 
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
         Left            =   420
         TabIndex        =   104
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   885
         TabIndex        =   103
         ToolTipText     =   "Número de Revisión"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   8400
         TabIndex        =   100
         Top             =   300
         Width           =   2220
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   3000
         TabIndex        =   11
         ToolTipText     =   "Doble Click Para Conmutar"
         Top             =   360
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   12015
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
         Height          =   270
         Left            =   11730
         TabIndex        =   39
         Top             =   360
         Width           =   165
      End
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   10800
         TabIndex        =   38
         Text            =   " "
         Top             =   360
         Width           =   945
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
         Left            =   11760
         TabIndex        =   20
         Top             =   840
         Width           =   165
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   10095
         TabIndex        =   19
         Text            =   " "
         Top             =   840
         Width           =   1680
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Index           =   0
         Left            =   3360
         TabIndex        =   22
         Top             =   120
         Width           =   5655
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descripción del Equipo Ensayado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   120
            Width           =   5175
         End
      End
      Begin VB.Label Label7 
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
         Index           =   0
         Left            =   9960
         TabIndex        =   40
         Top             =   435
         Width           =   810
      End
      Begin VB.Label Label92 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión30/05/16"
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
         Index           =   3
         Left            =   80
         TabIndex        =   37
         Top             =   360
         Width           =   1600
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
         Index           =   0
         Left            =   8640
         TabIndex        =   21
         Top             =   885
         Width           =   1380
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
         Left            =   -50
         TabIndex        =   18
         Top             =   885
         Width           =   795
      End
      Begin VB.Label Label95 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   2880
      End
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "55"
      Height          =   405
      Index           =   1
      Left            =   7200
      TabIndex        =   46
      Top             =   10845
      Width           =   945
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "55"
      Height          =   405
      Index           =   2
      Left            =   7200
      TabIndex        =   45
      Top             =   11250
      Width           =   945
   End
   Begin VB.Label Labe48 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "55"
      Height          =   405
      Index           =   3
      Left            =   7200
      TabIndex        =   44
      Top             =   11640
      Width           =   945
   End
   Begin VB.Label Label88 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   9960
      TabIndex        =   41
      Top             =   10800
      Width           =   3600
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
      Height          =   165
      Left            =   8040
      TabIndex        =   10
      Top             =   2565
      Width           =   1050
   End
End
Attribute VB_Name = "SEGELE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_COFUTIE$

Private Sub Combo1_Change()
  For I% = 0 To text41.UBound
     VALIDAR (I%)
  Next I%
End Sub

Private Sub Combo1_Click()
  'ACTUALIZA LOS COMBOS
  For I% = 0 To text41.UBound
     VALIDAR (I%)
  Next I%
End Sub


Private Sub Combo1_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub


Private Sub Command1_Click()
   '\\\OT-08-0936-OD-08/10/08///
    Call CIERRARCH("*.*")
    Unload Me
    '\\\OT-08-0936-OD-FIN///
End Sub
Private Sub Command2_Click()
   '\\\OT-08-0936-OD-08/10/08///
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
    If TRIM$(Label98) = "" Or XVALO(Text4) < 1 Then
      Call MENSERR(24, "Falta Informar Instrumento")
      GoTo 99
    End If

    'VALIDACION DE VEREDICTO GENERAL
    '
    If Text7 = "" Then
      Call MENSERR(24, "Falta Informar Veredicto General")
      GoTo 99
    End If
    '
    If TRIM$(UCase$(Text7)) <> "NC" And TRIM$(UCase$(Text7)) <> "A" Then
      Call COMUNI("Veredicto No Válido \Valores Posibles 'NC o A'")
      GoTo 99
    End If
    '
    If Label8 = "" Then
      Call MENSERR(24, "Falta Informar Inspector")
      GoTo 99
    End If
    '
    If TRIM$(Text1) = "" Then
      Call MENSERR(24, "Falta Informar Técnico")
      GoTo 99
    End If

    '
    If TRIM$(Combo1.TEXT) = "" Then
      Call MENSERR(24, "Falta Informar Clase")
      GoTo 99
    End If
    

    Screen.MousePointer = 11
    '
    '
    CONDI$ = "NumeroSerie = '" & NUSERIE$ & "'"
    CANTRE& = CantRegistros&("SEGELECT", CONDI$, "NOMESS")
    If CANTRE& = 0 Then
       REVIS% = 0
    Else
       REVIS% = CANTRE& 'NO LE SUMO UNO POR QUE ARRANCA DE 0 (EJ. SI DEVUELVE 0 REGISTROS VA A DAR DE ALTA EL NUMERO 0, SI DEVUELVE 1 REGISTRO
                        'VA A DAR DE ALTA EL 1, POR QUE EL REGISTRO 1 CORRESPONDE AL 0
    End If

    CIXI% = CODINT%(Label95)
    SQLX$ = "SELECT * FROM SEGELECT" '\\\OT-08-1064-OD-21/11/08-MEDITEA-AGREGO LA NUEVA REVISION
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & Text8 & "'"
    SQLX$ = SQLX$ + " AND REVISION = " & REVIS%
    Dim ENSA2TMP As ADODB.Recordset
    Set ENSA2TMP = New ADODB.Recordset
    ENSA2TMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With ENSA2TMP
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !NumeroSerie = TRIM$(Text8)
            !MODELO = CIXI%
            !Inspector = Left$(TRIM$(Label8), 32)
            !Comentarios = Text9
            !CorrFuga1 = XVALO(Text2(0))
            !CorrClase1 = XVALO(text41(0))
            !Veredicto1 = Left$(TRIM$(Text3(0)), 6)
            !CorrFuga2 = XVALO(Text2(1))
            !CorrClase2 = XVALO(text41(1))
            !Veredicto2 = Left$(TRIM$(Text3(1)), 6)
            !CorrFuga3 = XVALO(Text2(2))
            !CorrClase3 = XVALO(text41(2))
            !Veredicto3 = Left$(TRIM$(Text3(2)), 6)
            !CorrFuga4 = XVALO(Text2(3))
            !CorrClase4 = XVALO(text41(3))
            !Veredicto4 = Left$(TRIM$(Text3(3)), 6)
            !VerGeneral = Left$(TRIM$(Text7), 6)
            !Bloqueado = "SI"
            !Fech_Ensa = Text11
            !TECNICO = Left$(TRIM$(USERNAME$), 32) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-GRABACION DEL TECNIO
            !Revision = REVIS% ' GRABA CAMPO DE REVISION
            !TENSION = Left$(Label12, 32)
            !INSTRUMENTO = Left$(Label98, 32)
            !Cod_Instrumento = XVALO(Text4)
            !CLASE = TRIM$(Combo1.TEXT)
            !Frec_Alim = Mid$(TRIM$(Label1), 38)
          
            .Update
    End With
    '
    ENSA2TMP.Close
    Set ENSA2TMP = Nothing
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Ensayo de Seguridad Eléctrica" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Text1))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text7))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    'Call PRINTMEDI(1) '  0 IMPRESION SIMPLE --- UNO GRABACION
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(REVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("SEGE-MDT", FILTX$)
    '
    Image1.Visible = True
    Image3.Visible = False
    Command23.Enabled = True
    '
    Call BLANFORMU
    '
    Screen.MousePointer = 1
    '
99  Command2.Enabled = True
   '\\\OT-08-0936-OD-fin///
End Sub

Private Sub Command23_Click()
   NUSERIE$ = TRIM$(Text8)
   CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
   If CANTRE& < 1 Then
      Call MENSERR(24, "Número de Serie Incorrecto")
      Exit Sub
   End If
'   Call PRINTMEDI(0) '  0 IMPRESION SIMPLE --- UNO GRABACION
    ULTREVIS% = XVALO(Label13)
    CIXI% = CODINT%(TRIM$(Label95))
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(ULTREVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("SEGE-MDT", FILTX$)
    '
End Sub
Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPFET.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
   '\\\OT-08-0936-OD-08/10/08///
   Call Text8_DblClick
   '\\\OT-08-0936-OD-FIN///
End Sub


Private Sub Command6_Click()
   '\\\OT-08-0936-OD-08/10/08///
    If Image3.Visible = False Then
      Call CommSel("SELFECHA", Command6, Text11)
    Else
      Command4.SetFocus
    End If
    '\\\OT-08-0936-OD-FIN///
End Sub

Sub Command7_Click()
       Command7.Enabled = False
       If DERACCE%("INSTRUME", "Carga de Instrumentos") < 2 Then GoTo 99
       Call TRALOCAL("INSTRUME", "")
12     VTB% = VENTABU%("INSTRUME/NC/TITUPAN=TABLA DE INSTRUMENTOS")
       If VTB% < 0 Or ULTREG&("!INSTRUME") < 1 Then
         Exit Sub
       End If
       '
       For YX& = 1 To ULTREG&("!INSTRUME")
         YY$ = REGLEIDO$("!INSTRUME", YX&)
         CI% = Asc(Mid$(YY$, 1, 1))
         DESCRIPX$ = TRIM$(Mid$(YY$, 2, 32))
         If CI% > 0 Then
           If DESCRIPX$ = "" Then
             Call COMUNI("Falta Descripción de Instruemento en Fila: " & YX&)
             GoTo 12
           End If
         End If
         '
         If DESCRIPX$ <> "" Then
           If CI% < 1 Then
             Call COMUNI("Falta Código de Instruemento en Fila: " & YX&)
             GoTo 12
           End If
         End If
       Next YX&
       '
       Call TRACENTRAL("INSTRUME", "")
       '
99     Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Text4 = CODINSTRU%
End Sub

Private Sub Command8_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Form_Load()
    '\\\OT-08-0936-OD-06/10/08///
    Call ABREORFAB
    Call BLANFORMU
    FF% = HOY(HO$)
    Text11.TEXT = HO$
    Label8 = USERNAME$
    Label12 = "242 V"
    Combo1.Clear
    Combo1.AddItem "Clase I"
    Combo1.AddItem "Clase II"
    Combo1.TEXT = "Clase I"
    
'    Combo2.Clear
'    Combo2.AddItem "50 HZ"
'    Combo2.AddItem "60 HZ"
'    Combo2.TEXT = "50 HZ"
   
    '\\\OT-08-0936-OD-FIN///
End Sub
Private Sub Frame5_DragDrop(Source As CONTROL, X As Single, Y As Single)
End Sub
Private Sub Image1_DblClick()
   If DERACCE("MODENS72", "Modificacion  de Ensayo de Fuga a Tierra") > 1 Then
        Image1.Visible = False
        Image3.Visible = True
        Command2.Enabled = False
        Command23.Enabled = True '\\\OT-08-1064-OD-21/11/08-MEDITEA- HABILITA LA IMPRESORA
        If TRIM$(Text8) <> "" Then
          'PONGO LA FECHA DEL ULTIMO INFORME
          NUSERIE$ = TRIM$(Text8)
          CONDI$ = "NumeroSerie ='" & NUSERIE$ & "'"
          ULTIMOENSA% = XVALO(VALOBADA("SEGELECT", "MAX(REVISION)", CONDI$))
          CONDI$ = CONDI$ + " AND REVISION = " & ULTIMOENSA%
          Text11 = FECHATEX(CVINT(VALOBADA("SEGELECT", "Fech_Ensa", CONDI$, "NOMESS")))
        End If
   End If
End Sub
Private Sub Image3_DblClick()
   If DERACCE("MODENS72", "Modificacion  de Ensayo de Fuga a Tierra") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False
        'PONGO FECHA DEL DIA CUANDO VA A REALIZAR OTRA REVISION
        HOII% = HOY(HOS$)
        Text11 = HOS$
   End If
End Sub

Private Sub Label103_Change()
End Sub

Private Sub Label97_Click()

End Sub

Private Sub Label98_DblClick()
    '
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label98 = ECOARCH$("INSTRUME", Chr$(CODINSTRU%))
    Text4 = CODINSTRU%
    '

End Sub
Private Sub Ltota_Click(Index As Integer)
End Sub

Private Sub Text1_GotFocus()
  If Image3.Visible = True Then Command4.SetFocus '\\\OT-08-1066-OD-21/11/08-MEDITEA QUE NO SE PUEDA EDITAR SI TIENE EL CANDADO VISIBLE
End Sub
Private Sub Text11_DblClick()
   Call Command6_Click
End Sub
Private Sub Text11_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub
Private Sub Text2_Change(Index As Integer)
   Call VALIDAR(Index)
'   If XVALO(Text2(Index)) > 0 Then '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR PARA AGREGAR A O NC O LIMPIAR DE ACUERDO AL INGRESO EN TEXT2
'    If XVALO(Labe144(Index)) >= XVALO(Text2(Index)) Then
'      Text3(Index) = "A"
'    Else
'      Text3(Index) = "NC"
'    End If
'   Else
'    Text3(Index) = ""
'   End If
End Sub
Sub VALIDAR(I%)
    Text3(I%) = ""
    APROBADO% = 0
    If XVALO(Text2(I%)) > 0 And XVALO(text41(I%)) > 0 Then
      If XVALO(Labe144(I%)) >= XVALO(Text2(I%)) Then
        APROBADO% = 1
      Else
         APROBADO% = 4
      End If
      '
      VARI% = XVALO(Labe143(I%))
      If Combo1.TEXT = "Clase I" Then
        VARI% = XVALO(Labe142(I%))
      End If
      '
      If VARI% >= XVALO(text41(I%)) Then
        APROBADO% = APROBADO% + 1
      Else
         APROBADO% = 8
      End If
   End If
   If APROBADO% > 2 Then Text3(I%) = "NC"
   If APROBADO% = 2 Then Text3(I%) = "A"
End Sub

Private Sub Text4_Change()
     Label98 = ECOARCH$("INSTRUME", Chr$(XVALO(Text4)))
End Sub

Private Sub Text41_Change(Index As Integer)
    Call VALIDAR(Index)
'   VARI% = XVALO(Labe143(Index))
'   If Combo1.TEXT = "Clase I" Then
'      VARI% = XVALO(Labe142(Index))
'   End If
'   '
'   If XVALO(text41(Index)) > 0 Then '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR PARA AGREGAR A O NC O LIMPIAR DE ACUERDO AL INGRESO EN TEXT2
'    If VARI% >= XVALO(text41(Index)) Then
'      Text3(Index) = "A"
'    Else
'      Text3(Index) = "NC"
'    End If
'   Else
'    Text3(Index) = ""
'   End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   If Image3.Visible = True Then
      Command4.SetFocus
      Exit Sub
   End If
    
   Call PINTATEXT(Text2(Index))
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index < 3 Then
      Call PINTATEXT(Text2(Index + 1))
      'Text2(Index + 1).SetFocus
      
    Else
      Call PINTATEXT(Text2(0))
'      Text2(0).SetFocus
    End If
  End If

End Sub

Private Sub Text2_LostFocus(Index As Integer)
  VALOR1 = ROUND#(XVALO(Text2(Index)))
 
  If XVALO(VALOR1) < 10000 Then
    Text2(Index) = FORMATNUM(Text2(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    Text2(Index).SetFocus
  End If
End Sub
Private Sub Text3_Change(Index As Integer)
   Call VeredGeneral
End Sub
Private Sub Text3_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub text41_GotFocus(Index As Integer)
   If Image3.Visible = True Then
      Command4.SetFocus
      Exit Sub
   End If
    
   Call PINTATEXT(text41(Index))
End Sub

Private Sub text41_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index < 3 Then
      Call PINTATEXT(text41(Index + 1))
      
    Else
      Call PINTATEXT(text41(0))
    End If
  End If

End Sub

Private Sub text41_LostFocus(Index As Integer)
  VALOR1 = ROUND#(XVALO(text41(Index)))
 
  If XVALO(VALOR1) < 10000 Then
    text41(Index) = FORMATNUM(text41(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    text41(Index).SetFocus
  End If
  
End Sub

Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub
Private Sub Text8_Change()
   
   '\\\OT-08-0936-OD-08/10/08///
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Visible = True
   Image1.Enabled = False
   Command23.Enabled = True
   TXTRAZA_COFUTIE$ = ""
   Call LIMPIAINFORMACION
   '\\\OT-08-0936-OD-08/10/08///
End Sub
Private Sub Text8_DblClick()
   '\\\OT-08-0936-OD-08/10/08///
    Call SELECHO("INDINUSE")
    TTXX$ = VALACT1$("INDINUSE")
    CODI$ = TRIM$(Left$(VALACT2$("INDINUSE"), 16))
    If TTXX$ <> "" Then
       Text8 = TTXX$
      Call Text8_KeyPress(13)
    End If
 
   '\\\OT-08-0936-OD-08/10/08///
End Sub
Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label13) ' SI VIENE DE IMPRESION SIMPLE TOMA DEL LABEL
    If N% > 0 Then
        REVIS% = 0 ' SI VIENE DE GRABACION BUSCA EL  VALOR MAXIMO
        NUSERIE$ = TRIM$(Text8)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM SEGELECT WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        Dim MAXREVTMP As ADODB.Recordset
        'Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        Set MAXREVTMP = READSET(SQLZ$)
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    '
    TIPODOC$ = "Fuga a Tierra" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Text8) & "-" & REVIS%
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Text11) 'FECHA
    CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = Label88.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(3) = "REF-MAT1":  DOCPARAM$(3) = TRIM$(Text8.TEXT) 'NRO SERIE
    CODPARAM$(4) = "REF-MAT2":  DOCPARAM$(4) = UCase(TRIM$(Text7.TEXT))  'VEREDICTO
    CODPARAM$(5) = "REF-MAT3":  DOCPARAM$(5) = TRIM$(Label8)  'Inspector
    CODPARAM$(6) = "OBSERVA":  DOCPARAM$(6) = TRIM$(Text9.TEXT)  'Observaciones
    CODPARAM$(7) = "REF-MAT4":  DOCPARAM$(7) = TRIM$(Label12)  'TENSION
    CODPARAM$(8) = "REF-MAT5":  DOCPARAM$(8) = TRIM$(Label98)  'INSTRUMENTO
    CAPARDOC% = 8
    '
    CODTABU1$(1) = "MEDIDA-1"
    CODTABU1$(2) = "MEDIDA-2"
    CACOLTAB1% = 2
    '
    CAITAB1% = 0
    For U& = 1 To 8
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Text2(U& - 1))    'MEDIDA
        TETABU1$(2, CAITAB1%) = UCase(TRIM$(Text3(U& - 1)))   ' VEREDICTO
    Next U&
   
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FUENTIE"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    For I& = 1 To 4
        Text2(I& - 1) = ""
        Text3(I& - 1) = ""
        text41(I& - 1) = ""
    Next I&
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text4 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Text11 = FECHATEX$(HOY(HO$))
    'Label12 = "242 V"
'    Label98 = "I354"
    
End Sub
Sub LIMPIAINFORMACION()
    For I& = 1 To 4
        Text2(I& - 1) = ""
        Text3(I& - 1) = ""
    Next I&
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text7 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Label98 = ""
    RIGIDIEL07.Label98 = ""
    RIGIDIEL07.Text5 = ""

End Sub
Sub VeredGeneral()
  '
  CONTAD% = 0
  '
'  For I& = 0 To 3 '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR SI HAY UN NC VEREDICTO SEA NC
'    If UCase(TRIM$(Text3(I&))) = "NC" Then
'      Text7 = "NC"
'      Exit Sub
'    End If
'  Next I&
  '
  For I& = 0 To 3
    If UCase(TRIM$(Text3(I&))) = "A" Then
      CONTAD% = CONTAD% + 1
    End If
  Next I&
  '
'  VACIO% = 0
'  For I& = 0 To 3
'    If UCase(TRIM$(Text3(I&))) = "" Then
'      VACIO% = VACIO% + 1
'    End If
'  Next I&

  If CONTAD% = 4 Or CONTAD% + VACIO% = 4 Then 'VALIDAR SI TODOS TIENEN A  VEREDICTO SEA A
    Text7 = "A"
  ElseIf CONTAD% > 0 Then
    Text7 = "NC"
  Else
    Text7 = ""
  End If
  '
End Sub

Private Sub Text8_GotFocus()
'  Call Command4.SetFocus
'  Call Command4_Click
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   
           NUSERIE$ = TRIM$(Text8)
           CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
           If CANTRE& < 1 Then
               Call MENSERR(24, "Número de Serie Incorrecto")
               Exit Sub
           End If
           '
'           Label12 = "242 V"
        '   Label98 = "I354"
        TXTRAZA_COFUTIE$ = TXTRAZA_COFUTIE$ & "PUNTO 1 = " & Now & vbCrLf
           If InStr(NUSERIE$, "'") < 1 Then
               Screen.MousePointer = 11
               CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NumeroSerie ='" & NUSERIE$ & "'"))
               If CI% > 0 Then
                Label95 = CODEXT$(CI%)
                Label88 = DESCRIT0$(CI%)
               End If
               REVIS% = 0
               '
               'SI EL CODIGO COMIENZA CON 11 ES A 110 V ELSE 220 V.
               DATO$ = ""
               If Left$(TRIM$(Label95), 2) = "11" Then
                   DATO$ = TALLE$(CI%)
               Else
                   DATO$ = Color$(CI%)
               End If
               Label1.Caption = "Tensión de Alimentación y Frecuencia: " & DATO$
               '
               BLQUEO$ = ""
               SQLX$ = "SELECT TOP 1 *  FROM SEGELECT WITH(NOLOCK)"
               SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
               'SQLX$ = SQLX$ + " AND REVISION = " & ULTREVIS% & "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-AGREGO EL NUMERO DE REVISION COMO PARTE DE LA CONSULTA
               SQLX$ = SQLX$ + " order by REVISION DESC"
               Dim ENSA2TMP As ADODB.Recordset
               Set ENSA2TMP = READSET(SQLX$)
               With ENSA2TMP
                    If Not .EOF Then
                        Label13 = XVALO(!Revision)
                        CI% = !MODELO
                        Label95 = CODEXT$(CI%)
                        Label88 = DESCRIT0$(CI%)
                        Label8 = SAFETEXT$(!Inspector)
                        Text9 = SAFETEXT$(!Comentarios)
                        Text2(0) = FORMATNUM(XVALO(!CorrFuga1), "F7.2")
                        Text3(0) = SAFETEXT$(!Veredicto1)
                        Text2(1) = FORMATNUM(XVALO(!CorrFuga2), "F7.2")
                        Text3(1) = SAFETEXT$(!Veredicto2)
                        Text2(2) = FORMATNUM(XVALO(!CorrFuga3), "F7.2")
                        Text3(2) = SAFETEXT$(!Veredicto3)
                        Text2(3) = FORMATNUM(XVALO(!CorrFuga4), "F7.2")
                        Text3(3) = SAFETEXT$(!Veredicto4)
                        
                        text41(0) = FORMATNUM(XVALO(!CorrClase1), "F7.2")
                        text41(1) = FORMATNUM(XVALO(!CorrClase2), "F7.2")
                        text41(2) = FORMATNUM(XVALO(!CorrClase3), "F7.2")
                        text41(3) = FORMATNUM(XVALO(!CorrClase4), "F7.2")
                        
                        Text1 = SAFETEXT$(!TECNICO) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-CARGA EL TEXT DEL TECNIO
                        BLQUEO$ = SAFETEXT$(!Bloqueado)
                        Text7 = SAFETEXT$(!VerGeneral)
                        fecccha$ = SAFETEXT$(!Fech_Ensa) ' si no tiene fecha pone la de hoy
                        If fecccha$ <> "" Then fff% = CVINT(fecccha$): Text11 = FECHATEX(fff%)
                        'si tiene fecha le damos formato "dd/mm/aa"
                        If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Text11.TEXT = HO$
                        Label12 = SAFETEXT$(!TENSION): If Label12 = "" Then Label12 = "242 V"
                        
                        Text4 = XVALO(!Cod_Instrumento)
                        'Label98 = SAFETEXT$(!INSTRUMENTO)
                        Combo1.TEXT = SAFETEXT$(!CLASE)
                        'Combo2.TEXT = SAFETEXT$(!Frec_Alim)
                    End If
              End With
          ENSA2TMP.Close
          Set ENSA2TMP = Nothing
          Screen.MousePointer = 1

          End If
          '
          If BLQUEO$ = "SI" Then
                Image3.Visible = True
                Image1.Visible = False
                Image1.Enabled = True
                Command2.Enabled = False
          Else
                Command23.Enabled = False '\\\OT-08-1064-OD-21/11/08-MEIDTEA BLOQUEA LA IMPRESION
                Text1 = Label8
          End If
          'GRABO EL ARCHIVO DE TRAZABILIDAD
'        For I1% = 1 To 32767
'             ARCH$ = LUDAT$ & "TXTRAZA_COFUTIE-" & TRIM$(Str$(I1%)) & ".TXT"
'             If EXISTE%(ARCH$) < 1 Then
'               NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'               Print #NARCHI2%, TXTRAZA_COFUTIE$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'               Close #NARCHI2%
'               Exit For
'             End If
'        Next I1%
    End If
End Sub

Private Sub Text9_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub




