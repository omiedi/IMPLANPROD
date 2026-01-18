VERSION 5.00
Begin VB.Form COFUTIE07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Ensayo de Fuga a Tierra"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11190
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
   ScaleHeight     =   8505
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   2520
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   7575
      TabIndex        =   0
      Text            =   " "
      Top             =   670
      Width           =   1680
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
      Left            =   9240
      TabIndex        =   1
      Top             =   670
      Width           =   165
   End
   Begin VB.TextBox Text9 
      Height          =   750
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1830
      Width           =   5730
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   9470
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   705
   End
   Begin VB.TextBox Text11 
      Height          =   300
      Left            =   9000
      TabIndex        =   2
      Text            =   " "
      Top             =   280
      Width           =   945
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
      Height          =   270
      Left            =   9930
      TabIndex        =   3
      Top             =   280
      Width           =   165
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
      Left            =   10370
      ScaleHeight     =   8655
      ScaleWidth      =   870
      TabIndex        =   30
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command7 
         Height          =   650
         Left            =   105
         Picture         =   "COFUTIE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Tabla de Instrumentos"
         Top             =   3045
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Height          =   650
         Left            =   105
         Picture         =   "COFUTIE07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Datos Complementarios"
         Top             =   5880
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   105
         Picture         =   "COFUTIE07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Visualiza e Imprime Informe de  Fuga a Tierra"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Height          =   650
         Left            =   105
         Picture         =   "COFUTIE07.frx":08FE
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   650
         Left            =   120
         Picture         =   "COFUTIE07.frx":0D40
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Graba e Imprime Informe"
         Top             =   6720
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
         TabIndex        =   31
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
            TabIndex        =   32
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Height          =   630
         Left            =   105
         Picture         =   "COFUTIE07.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "COFUTIE07.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "COFUTIE07.frx":149E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7970
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Corriente de Fuga a Tierra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   240
      TabIndex        =   28
      Top             =   2880
      Width           =   9960
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   6540
         TabIndex        =   14
         Text            =   " "
         Top             =   2970
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   6540
         TabIndex        =   16
         Text            =   " "
         Top             =   3285
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   6540
         TabIndex        =   18
         Text            =   " "
         Top             =   3600
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   7
         Left            =   6540
         TabIndex        =   20
         Text            =   " "
         Top             =   3915
         Width           =   1050
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   4
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   2970
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   5
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   " "
         Top             =   3285
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   6
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   " "
         Top             =   3600
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   7
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   3915
         Width           =   1130
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   6540
         TabIndex        =   6
         Text            =   " "
         Top             =   1785
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   6540
         TabIndex        =   8
         Text            =   " "
         Top             =   2085
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   6540
         TabIndex        =   10
         Text            =   " "
         Top             =   2370
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   6540
         TabIndex        =   12
         Text            =   " "
         Top             =   2670
         Width           =   1050
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   0
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   " "
         Top             =   1785
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   1
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   " "
         Top             =   2085
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   2
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   2370
         Width           =   1130
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   3
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   2670
         Width           =   1130
      End
      Begin VB.TextBox text19 
         Height          =   405
         Index           =   1
         Left            =   11760
         TabIndex        =   59
         Top             =   2085
         Width           =   795
      End
      Begin VB.TextBox text19 
         Height          =   405
         Index           =   0
         Left            =   11520
         TabIndex        =   58
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota 5:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6240
         TabIndex        =   121
         Top             =   4710
         Width           =   975
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ON: Neutro abierto"
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
         Left            =   1200
         TabIndex        =   134
         Top             =   5040
         Width           =   5070
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota 3:"
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
         Left            =   240
         TabIndex        =   133
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Height          =   220
         Left            =   3045
         TabIndex        =   132
         ToolTipText     =   "Doble Click Para Conmutar"
         Top             =   520
         Width           =   840
      End
      Begin VB.Label Label41 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R: Reversa"
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
         Left            =   7200
         TabIndex        =   125
         Top             =   5040
         Width           =   2600
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N: Normal"
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
         Left            =   7200
         TabIndex        =   124
         Top             =   4750
         Width           =   2600
      End
      Begin VB.Label Label39 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0: Interruptor apagado"
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
         Left            =   7200
         TabIndex        =   123
         Top             =   4470
         Width           =   2600
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1: Interruptor abierto"
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
         Left            =   7200
         TabIndex        =   122
         Top             =   4200
         Width           =   2600
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota 4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6240
         TabIndex        =   120
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPD: Condición de primer defecto"
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
         Left            =   1200
         TabIndex        =   119
         Top             =   4760
         Width           =   5055
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CN: Condición Normal"
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
         Left            =   1200
         TabIndex        =   118
         Top             =   4480
         Width           =   5055
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER: Corriente de fuga a tierra"
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
         Left            =   1200
         TabIndex        =   117
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   116
         Top             =   4470
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota 1:"
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
         Left            =   240
         TabIndex        =   115
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   5
         Left            =   1200
         TabIndex        =   114
         Top             =   3285
         Width           =   1060
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   4
         Left            =   1200
         TabIndex        =   113
         Top             =   2970
         Width           =   1060
      End
      Begin VB.Label Labe46 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   11
         Left            =   1200
         TabIndex        =   112
         Top             =   3600
         Width           =   1060
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   7
         Left            =   1200
         TabIndex        =   111
         Top             =   3915
         Width           =   1060
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   300
         Index           =   5
         Left            =   5475
         TabIndex        =   110
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         Height          =   300
         Index           =   4
         Left            =   5475
         TabIndex        =   109
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   300
         Index           =   6
         Left            =   5475
         TabIndex        =   108
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         Height          =   300
         Index           =   7
         Left            =   5475
         TabIndex        =   107
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         Height          =   300
         Index           =   5
         Left            =   7605
         TabIndex        =   106
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         Height          =   300
         Index           =   4
         Left            =   7605
         TabIndex        =   105
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         Height          =   300
         Index           =   6
         Left            =   7605
         TabIndex        =   104
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         Height          =   300
         Index           =   7
         Left            =   7605
         TabIndex        =   103
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         Height          =   300
         Index           =   7
         Left            =   240
         TabIndex        =   102
         Top             =   3915
         Width           =   945
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         Height          =   300
         Index           =   6
         Left            =   240
         TabIndex        =   101
         Top             =   3600
         Width           =   950
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   300
         Index           =   5
         Left            =   240
         TabIndex        =   100
         Top             =   3285
         Width           =   950
      End
      Begin VB.Label Labe137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   300
         Index           =   4
         Left            =   240
         TabIndex        =   99
         Top             =   2970
         Width           =   950
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPD"
         Height          =   300
         Index           =   5
         Left            =   2280
         TabIndex        =   98
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPD"
         Height          =   300
         Index           =   4
         Left            =   2280
         TabIndex        =   97
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ON"
         Height          =   300
         Index           =   5
         Left            =   3345
         TabIndex        =   96
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPD"
         Height          =   300
         Index           =   7
         Left            =   2280
         TabIndex        =   95
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Index           =   5
         Left            =   4410
         TabIndex        =   94
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CPD"
         Height          =   300
         Index           =   6
         Left            =   2280
         TabIndex        =   93
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ON"
         Height          =   300
         Index           =   6
         Left            =   3345
         TabIndex        =   92
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ON"
         Height          =   300
         Index           =   4
         Left            =   3345
         TabIndex        =   91
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Index           =   4
         Left            =   4410
         TabIndex        =   90
         Top             =   2970
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Index           =   6
         Left            =   4410
         TabIndex        =   89
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Index           =   7
         Left            =   4410
         TabIndex        =   88
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ON"
         Height          =   300
         Index           =   7
         Left            =   3345
         TabIndex        =   87
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   86
         Top             =   2085
         Width           =   1060
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   85
         Top             =   1785
         Width           =   1060
      End
      Begin VB.Label Labe138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   84
         Top             =   2370
         Width           =   1060
      End
      Begin VB.Label label138 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER"
         Height          =   300
         Index           =   3
         Left            =   1200
         TabIndex        =   83
         Top             =   2670
         Width           =   1060
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   300
         Index           =   1
         Left            =   5475
         TabIndex        =   82
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         Height          =   300
         Index           =   0
         Left            =   5475
         TabIndex        =   81
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         Height          =   300
         Index           =   2
         Left            =   5475
         TabIndex        =   80
         Top             =   2370
         Width           =   1050
      End
      Begin VB.Label Labe143 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         Height          =   300
         Index           =   3
         Left            =   5475
         TabIndex        =   79
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         Height          =   300
         Index           =   1
         Left            =   7605
         TabIndex        =   78
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         Height          =   300
         Index           =   0
         Left            =   7605
         TabIndex        =   77
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         Height          =   300
         Index           =   2
         Left            =   7605
         TabIndex        =   76
         Top             =   2370
         Width           =   1050
      End
      Begin VB.Label Labe144 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         Height          =   300
         Index           =   3
         Left            =   7605
         TabIndex        =   75
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   74
         Top             =   2670
         Width           =   950
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   73
         Top             =   2370
         Width           =   950
      End
      Begin VB.Label Labe137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   295
         Index           =   1
         Left            =   240
         TabIndex        =   72
         Top             =   2085
         Width           =   950
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frecuencia de Alimentación: 50 Hz"
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
         Left            =   5040
         TabIndex        =   71
         Top             =   480
         Width           =   4770
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tensión de Alimentación: "
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
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   5010
      End
      Begin VB.Label Labe48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   405
         Index           =   3
         Left            =   10965
         TabIndex        =   57
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   405
         Index           =   2
         Left            =   10965
         TabIndex        =   56
         Top             =   2490
         Width           =   945
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   405
         Index           =   1
         Left            =   10965
         TabIndex        =   55
         Top             =   2085
         Width           =   945
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   300
         Index           =   3
         Left            =   3345
         TabIndex        =   54
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Index           =   3
         Left            =   4410
         TabIndex        =   53
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Index           =   2
         Left            =   4410
         TabIndex        =   52
         Top             =   2370
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Index           =   0
         Left            =   4410
         TabIndex        =   51
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   300
         Index           =   0
         Left            =   3345
         TabIndex        =   50
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   300
         Index           =   2
         Left            =   3345
         TabIndex        =   49
         Top             =   2370
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CN"
         Height          =   300
         Index           =   2
         Left            =   2280
         TabIndex        =   48
         Top             =   2370
         Width           =   1050
      End
      Begin VB.Label Labe142 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Index           =   1
         Left            =   4410
         TabIndex        =   47
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CN"
         Height          =   300
         Index           =   3
         Left            =   2280
         TabIndex        =   46
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Labe141 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   300
         Index           =   1
         Left            =   3345
         TabIndex        =   45
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CN"
         Height          =   300
         Index           =   0
         Left            =   2280
         TabIndex        =   44
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Labe140 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CN"
         Height          =   300
         Index           =   1
         Left            =   2280
         TabIndex        =   43
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Label137 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   1785
         Width           =   950
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "            Veredicto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   8670
         TabIndex        =   41
         Top             =   840
         Width           =   1130
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corriente de fuga máxima per- mitida (uA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   7605
         TabIndex        =   40
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corriente de fuga máxima medida(uA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   6540
         TabIndex        =   39
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Polaridad (Nota 5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   5460
         TabIndex        =   38
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado del interruptor  (Nota 4)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   4410
         TabIndex        =   37
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    CPD    (Nota 3 )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   3360
         TabIndex        =   36
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condición  de Ensayo (Nota 2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   2280
         TabIndex        =   35
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "             Fila N."
         Height          =   885
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Corriente        de Fuga     (Nota 1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1200
         TabIndex        =   33
         Top             =   840
         Width           =   1060
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   9570
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   9760
      TabIndex        =   131
      ToolTipText     =   "Número de Revisión"
      Top             =   720
      Width           =   360
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
      Left            =   9360
      TabIndex        =   130
      Top             =   780
      Width           =   300
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   360
      TabIndex        =   129
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   128
      Top             =   1680
      Width           =   2520
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3840
      Picture         =   "COFUTIE07.frx":33C0
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3840
      Picture         =   "COFUTIE07.frx":36CA
      Top             =   2040
      Width           =   480
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
      Left            =   8160
      TabIndex        =   127
      Top             =   360
      Width           =   810
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
      Left            =   7200
      TabIndex        =   126
      Top             =   2700
      Width           =   2220
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
      Left            =   6120
      TabIndex        =   69
      Top             =   720
      Width           =   1380
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
      Height          =   405
      Left            =   6500
      TabIndex        =   68
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label95 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   67
      Top             =   620
      Width           =   2880
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
      Left            =   960
      TabIndex        =   66
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label88 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   430
      Left            =   1080
      TabIndex        =   65
      Top             =   980
      Width           =   3600
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
      Left            =   360
      TabIndex        =   64
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label98 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7590
      TabIndex        =   63
      ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
      Top             =   1035
      Width           =   2535
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
      Left            =   3240
      TabIndex        =   62
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label100 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analizador de corriente de fuga"
      Height          =   315
      Left            =   6120
      TabIndex        =   61
      Top             =   1460
      Width           =   4005
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Informe de Ensayo de Fuga a Tierrra"
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
      Left            =   120
      TabIndex        =   60
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "COFUTIE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_COFUTIE$

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
    If TRIM$(Label98) = "" Then
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

    'VALIDACIONES DE INFORMACION CARGADA EN FRM DE RIGIDEZ DIELECTRICA
    '
    If TRIM$(RIGIDIEL07.Text2) = "" Then  ' AHORA ES FIJO SEGUN FABIAN CORNEJ0
      Call COMUNI("Falta Tensión Utilizada de Rigidez Dielectrica")
      RIGIDIEL07.Show 1
      GoTo 99
    End If
    '
    If XVALO(RIGIDIEL07.Text3) < 1 Then  ' AHORA ES FIJO SEGUN FABIAN CORNEJ0
      Call COMUNI("Tiempo en Segundos No Válido")
      RIGIDIEL07.Show 1
      GoTo 99
    End If

    If TRIM$(RIGIDIEL07.Label98) = "" Then
      Call COMUNI("Falta Instrumento de Rigidez Dielectrica")
      RIGIDIEL07.Show 1
      GoTo 99
    End If
    If TRIM$(RIGIDIEL07.Text5) = "" Then
      Call COMUNI("Falta Veredicto de Rigidez Dielectrica")
      RIGIDIEL07.Show 1
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
'    SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR072 WITH(NOLOCK) " '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'    SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
''    Dim ENSA072TMP As ADODB.Recordset
''    Set ENSA072TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'    Set ENSA072TMP = READSET(SQLZ$)
'    ULTREVIS% = XVALO(ENSA072TMP!MAXREVI)
'    ENSA072TMP.Close
'    Set ENSA072TMP = Nothing
'    '
'    SQLX$ = "SELECT REVISION FROM ENSAR072 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- VALIDAR EL  VALOR DE CAMPO REVISION
'    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
'    '
''    Dim ENSA072TMP1 As ADODB.Recordset '\\\OT-08-1064-OD-21/11/08-MEDITEA-BUSCO SI ES LA PRIMERA REVISION ,CASO CONTRARIO ULTIMA REVISION + 1
''    Set ENSA072TMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    Set ENSA072TMP1 = READSET(SQLX$)
'    With ENSA072TMP1
'     If Not .EOF Then
'       REVIS% = 1 + ULTREVIS%
'     Else
'       REVIS% = 0 'SI ES EL PRIMERO EL VALOR VA A SER 0
'     End If
'    End With
'    ENSA072TMP1.Close
'    Set ENSA072TMP1 = Nothing
    '
    CONDI$ = "NumeroSerie = '" & NUSERIE$ & "'"
    CANTRE& = CantRegistros&("ENSAR072", CONDI$, "NOMESS")
    If CANTRE& = 0 Then
       REVIS% = 0
    Else
       REVIS% = CANTRE& 'NO LE SUMO UNO POR QUE ARRANCA DE 0 (EJ. SI DEVUELVE 0 REGISTROS VA A DAR DE ALTA EL NUMERO 0, SI DEVUELVE 1 REGISTRO
                        'VA A DAR DE ALTA EL 1, POR QUE EL REGISTRO 1 CORRESPONDE AL 0
    End If

    CIXI% = CODINT%(Label95)
    SQLX$ = "SELECT * FROM ENSAR072" '\\\OT-08-1064-OD-21/11/08-MEDITEA-AGREGO LA NUEVA REVISION
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
            !Veredicto1 = Left$(TRIM$(Text3(0)), 6)
            !CorrFuga2 = XVALO(Text2(1))
            !Veredicto2 = Left$(TRIM$(Text3(1)), 6)
            !CorrFuga3 = XVALO(Text2(2))
            !Veredicto3 = Left$(TRIM$(Text3(2)), 6)
            !CorrFuga4 = XVALO(Text2(3))
            !Veredicto4 = Left$(TRIM$(Text3(3)), 6)
            !CorrFuga5 = XVALO(Text2(4))
            !Veredicto5 = Left$(TRIM$(Text3(4)), 6)
            !CorrFuga6 = XVALO(Text2(5))
            !Veredicto6 = Left$(TRIM$(Text3(5)), 6)
            !CorrFuga7 = XVALO(Text2(6))
            !Veredicto7 = Left$(TRIM$(Text3(6)), 6)
            !CorrFuga8 = XVALO(Text2(7))
            !Veredicto8 = Left$(TRIM$(Text3(7)), 6)
            !VerGeneral = Left$(TRIM$(Text7), 6)
            !Bloqueado = "SI"
            !Fech_Ensa = Text11
            !TECNICO = Left$(TRIM$(Text1), 32) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-GRABACION DEL TECNIO
            !Revision = REVIS% ' GRABA CAMPO DE REVISION
            !TENSION = Left$(Label12, 32)
            !INSTRUMENTO = Left$(Label98, 32)
            !RIDIEINSTRU = Left$(RIGIDIEL07.Label98, 32)
            !RiDieTension = Left$(RIGIDIEL07.Text2, 32)
            !RIDIETIEMPO = Left$(RIGIDIEL07.Text3, 16)
            !RIDIEVEREDICTO = Left$(RIGIDIEL07.Text5, 6)
'            !RIDIEOBSERVA = RIGIDIEL07.Text1
           
            .Update
    End With
    '
    ENSA2TMP.Close
    Set ENSA2TMP = Nothing
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Corriente de Fuga a Tierra" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Text1))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text7))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    'Call PRINTMEDI(1) '  0 IMPRESION SIMPLE --- UNO GRABACION
    FILTX$ = NUSERIE$ & ";" & TRIM$(Str$(REVIS%)) & ";" & TRIM$(Str$(CIXI%))
    Call STDFORM("ENFT-MDT", FILTX$)
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
    Call STDFORM("ENFT-MDT", FILTX$)
    '
End Sub
Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPFET.Show 1
   Command26.Enabled = True
End Sub
Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If Image3.Visible = True Then
      Call COMUNI("Este Ensayo se Encuentra Cerrado\Si Desea Ver Datos de Rigidez Dieléctrica\Realice Vista Previa Desde el Boton Impresora")
      GoTo 99
   End If
   '
   NUSERIE$ = TRIM$(Text8)
   CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
   If CANTRE& < 1 Then
      Call MENSERR(24, "Número de Serie Incorrecto")
      GoTo 99
   End If
  
    SQLX$ = "SELECT * FROM ENSAR072"
    NUSERIE$ = TRIM$(Text8)
    ULTREVIS% = XVALO(Label13)
    CONDI$ = CONDI$ + "NumeroSerie ='" & NUSERIE$ & "'"
    CONDI$ = CONDI$ + " AND REVISION = " & ULTREVIS% & ""
    If CantRegistros("ENSAR072", CONDI$, "NOMESS") > 0 Then
         RIGIDIEL07.Label98 = VALOBADA("ENSAR072", "RIDIEINSTRU", CONDI$, "NOMESS")
         'RIGIDIEL07.Text2 = VALOBADA("ENSAR072", "RiDieTension", CONDI$, "NOMESS")
         'RIGIDIEL07.Text3 = VALOBADA("ENSAR072", "RIDIETIEMPO", CONDI$, "NOMESS")
         RIGIDIEL07.Text5 = VALOBADA("ENSAR072", "RIDIEVEREDICTO", CONDI$, "NOMESS")
    End If

   RIGIDIEL07.Show 1
99 Command3.Enabled = True
  
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

Private Sub Form_Load()
    '\\\OT-08-0936-OD-06/10/08///
    Call ABREORFAB
    Call BLANFORMU
    FF% = HOY(HO$)
    Text11.TEXT = HO$
    Label8 = USERNAME$
    Label12 = "242 V"
    '\\\OT-08-0936-OD-FIN///
End Sub
Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)
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
          ULTIMOENSA% = XVALO(VALOBADA("ENSAR072", "MAX(REVISION)", CONDI$))
          CONDI$ = CONDI$ + " AND REVISION = " & ULTIMOENSA%
          Text11 = FECHATEX(CVINT(VALOBADA("ENSAR072", "Fech_Ensa", CONDI$, "NOMESS")))
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

Private Sub Label12_DblClick()
    'CONMUTA DE 242 A 121
    If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
    NUSERIE$ = TRIM$(Text8)
    If NUSERIE$ <> "" Then
         If TRIM$(Label12) = "242 V" Then
            Label12 = "121 V"
         Else
            Label12 = "242 V"
         End If
    End If
   '
End Sub

Private Sub Label98_DblClick()
    '
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label98 = ECOARCH$("INSTRUME", Chr$(CODINSTRU%))
    '
   'CONMUTA DE I354 - I304 - I332
   '
'   CANCELADO POR LISTA DE SELECCION
'   If Image3.Visible = True Then Exit Sub ' SI ESTA BLOQUEADO NO CAMBIA
'   NUSERIE$ = TRIM$(Text8)
'   If NUSERIE$ <> "" Then
'        If TRIM$(Label98) = "I354" Then
'           Label98 = "I304"
'        ElseIf TRIM$(Label98) = "I304" Then
'           Label98 = "I332"
'        ElseIf TRIM$(Label98) = "I332" Then
'           Label98 = "I378"
'        ElseIf TRIM$(Label98) = "I378" Then
'           Label98 = "I379"
'        ElseIf TRIM$(Label98) = "I379" Then
'           Label98 = "I380"
'        Else
'           Label98 = "I354"
'        End If
'
'   End If
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
   If XVALO(Text2(Index)) > 0 Then '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR PARA AGREGAR A O NC O LIMPIAR DE ACUERDO AL INGRESO EN TEXT2
    If XVALO(Labe144(Index)) >= XVALO(Text2(Index)) Then
      Text3(Index) = "A"
    Else
      Text3(Index) = "NC"
    End If
   Else
    Text3(Index) = ""
   End If
End Sub
Private Sub Text2_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index < 7 Then
      Text2(Index + 1).SetFocus
    Else
      Text2(0).SetFocus
    End If
  End If

End Sub

Private Sub Text2_LostFocus(Index As Integer)
  VALOR1 = ROUND#(XVALO(Text2(Index)))
 
  If XVALO(VALOR1) < 10000 Then
    Text2(Index) = FORMATNUM(Text2(Index), "F4.0")
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
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR072 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
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
    FORIPRE$ = TRIM$(Control$("", "FUENTIE"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    For I& = 1 To 8
        Text2(I& - 1) = ""
        Text3(I& - 1) = ""
    Next I&
    Text1 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNIO
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Label13 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Text11 = FECHATEX$(HOY(HO$))
    Label12 = "242 V"
'    Label98 = "I354"
    RIGIDIEL07.Label98 = ""
    RIGIDIEL07.Text5 = ""
    
End Sub
Sub LIMPIAINFORMACION()
    For I& = 1 To 8
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
  For I& = 0 To 7 '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR SI HAY UN NC VEREDICTO SEA NC
    If UCase(TRIM$(Text3(I&))) = "NC" Then
      Text7 = "NC"
      Exit Sub
    End If
  Next I&
  '
  For I& = 0 To 7
    If UCase(TRIM$(Text3(I&))) = "A" Then
      CONTAD% = CONTAD% + 1
    End If
  Next I&
  '
  VACIO% = 0
  For I& = 0 To 7
    If UCase(TRIM$(Text3(I&))) = "" Then
      VACIO% = VACIO% + 1
    End If
  Next I&

  If CONTAD% = 8 Or CONTAD% + VACIO% = 8 Then 'VALIDAR SI TODOS TIENEN A  VEREDICTO SEA A
    Text7 = "A"
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
           Label12 = "242 V"
        '   Label98 = "I354"
        TXTRAZA_COFUTIE$ = TXTRAZA_COFUTIE$ & "PUNTO 1 = " & Now & vbCrLf
           If InStr(NUSERIE$, "'") < 1 Then
               Screen.MousePointer = 11
               CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NumeroSerie ='" & NUSERIE$ & "'"))
               If CI% > 0 Then
                Label95 = CODEXT$(CI%)
                Label88 = DESCRIT0$(CI%)
               End If
        TXTRAZA_COFUTIE$ = TXTRAZA_COFUTIE$ & "PUNTO 2 = " & Now & vbCrLf
               REVIS% = 0
'               SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR072 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'               SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'               Dim ENSA072TMP As ADODB.Recordset
'               'Set ENSA072TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'               Set ENSA072TMP = READSET(SQLZ$)
'               ULTREVIS% = XVALO(ENSA072TMP!MAXREVI)
'               ENSA072TMP.Close
'               Set ENSA072TMP = Nothing
'               Label13 = ULTREVIS%
        TXTRAZA_COFUTIE$ = TXTRAZA_COFUTIE$ & "PUNTO 3 = " & Now & vbCrLf
              
               BLQUEO$ = ""
               SQLX$ = "SELECT TOP 1 *  FROM ENSAR072 WITH(NOLOCK)"
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
                        Text2(0) = FORMATNUM(XVALO(!CorrFuga1), "F4.0")
                        Text3(0) = SAFETEXT$(!Veredicto1)
                        Text2(1) = FORMATNUM(XVALO(!CorrFuga2), "F4.0")
                        Text3(1) = SAFETEXT$(!Veredicto2)
                        Text2(2) = FORMATNUM(XVALO(!CorrFuga3), "F4.0")
                        Text3(2) = SAFETEXT$(!Veredicto3)
                        Text2(3) = FORMATNUM(XVALO(!CorrFuga4), "F4.0")
                        Text3(3) = SAFETEXT$(!Veredicto4)
                        Text2(4) = FORMATNUM(XVALO(!CorrFuga5), "F4.0")
                        Text3(4) = SAFETEXT$(!Veredicto5)
                        Text2(5) = FORMATNUM(XVALO(!CorrFuga6), "F4.0")
                        Text3(5) = SAFETEXT$(!Veredicto6)
                        Text2(6) = FORMATNUM(XVALO(!CorrFuga7), "F4.0")
                        Text3(6) = SAFETEXT$(!Veredicto7)
                        Text2(7) = FORMATNUM(XVALO(!CorrFuga8), "F4.0")
                        Text3(7) = SAFETEXT$(!Veredicto8)
                        Text1 = SAFETEXT$(!TECNICO) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-CARGA EL TEXT DEL TECNIO
                        BLQUEO$ = SAFETEXT$(!Bloqueado)
                        Text7 = SAFETEXT$(!VerGeneral)
                        fecccha$ = SAFETEXT$(!Fech_Ensa) ' si no tiene fecha pone la de hoy
                        If fecccha$ <> "" Then fff% = CVINT(fecccha$): Text11 = FECHATEX(fff%)
                        'si tiene fecha le damos formato "dd/mm/aa"
                        If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Text11.TEXT = HO$
                        Label12 = SAFETEXT$(!TENSION): If Label12 = "" Then Label12 = "242 V"
                        Label98 = SAFETEXT$(!INSTRUMENTO)
                        RIGIDIEL07.Label98 = SAFETEXT$(!RIDIEINSTRU)
                        RIGIDIEL07.Text5 = SAFETEXT$(!RIDIEVEREDICTO)
                    End If
              End With
          ENSA2TMP.Close
          Set ENSA2TMP = Nothing
          Screen.MousePointer = 1

          End If
          TXTRAZA_COFUTIE$ = TXTRAZA_COFUTIE$ & "PUNTO 4 = " & Now & vbCrLf
        
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




