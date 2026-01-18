VERSION 5.00
Begin VB.Form REGICON07 
   BackColor       =   &H00F0F0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Control Final"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Test"
      Height          =   255
      Left            =   5040
      TabIndex        =   218
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Instrumentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   217
      Top             =   6300
      Width           =   3615
      Begin VB.CommandButton Command11 
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
         Height          =   300
         Index           =   3
         Left            =   3280
         TabIndex        =   240
         ToolTipText     =   "Borra El Instrumento"
         Top             =   1275
         Width           =   270
      End
      Begin VB.CommandButton Command11 
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
         Height          =   300
         Index           =   2
         Left            =   3280
         TabIndex        =   239
         ToolTipText     =   "Borra El Instrumento"
         Top             =   960
         Width           =   270
      End
      Begin VB.CommandButton Command11 
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
         Height          =   300
         Index           =   1
         Left            =   3280
         TabIndex        =   238
         ToolTipText     =   "Borra El Instrumento"
         Top             =   645
         Width           =   270
      End
      Begin VB.CommandButton Command11 
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
         Height          =   300
         Index           =   0
         Left            =   3280
         TabIndex        =   237
         ToolTipText     =   "Borra El Instrumento"
         Top             =   330
         Width           =   270
      End
      Begin VB.Label Label109 
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
         Index           =   3
         Left            =   105
         TabIndex        =   236
         ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
         Top             =   1305
         Width           =   3165
      End
      Begin VB.Label Label109 
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
         Index           =   2
         Left            =   105
         TabIndex        =   235
         ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
         Top             =   990
         Width           =   3165
      End
      Begin VB.Label Label109 
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
         Index           =   1
         Left            =   105
         TabIndex        =   234
         ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
         Top             =   675
         Width           =   3165
      End
      Begin VB.Label Label109 
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
         Index           =   0
         Left            =   105
         TabIndex        =   233
         ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
         Top             =   360
         Width           =   3165
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   960
      Left            =   5640
      ScaleHeight     =   900
      ScaleWidth      =   6240
      TabIndex        =   50
      Top             =   0
      Width           =   6300
      Begin VB.CommandButton AGenRep 
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
         Left            =   1540
         Picture         =   "REGICON07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   203
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   120
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
         Left            =   4750
         Picture         =   "REGICON07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   2180
         Picture         =   "REGICON07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Imprime Registro de Control Final"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command26 
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
         Left            =   2820
         Picture         =   "REGICON07.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   120
         Width           =   650
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   0
         TabIndex        =   197
         Top             =   720
         Width           =   280
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
         Left            =   3470
         Picture         =   "REGICON07.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Graba e Imprime Registro"
         Top             =   120
         Width           =   650
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   960
         Left            =   195
         TabIndex        =   196
         Top             =   840
         Width           =   280
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   840
         TabIndex        =   194
         Top             =   120
         Width           =   900
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   195
            Top             =   0
            Width           =   12000
         End
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
         Left            =   4110
         Picture         =   "REGICON07.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   650
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
         Left            =   5400
         Picture         =   "REGICON07.frx":16D4
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   120
         Picture         =   "REGICON07.frx":181E
         Stretch         =   -1  'True
         Top             =   330
         Width           =   1170
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00F0F0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7035
      Left            =   3885
      TabIndex        =   64
      Top             =   960
      Width           =   7815
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0F0C0&
         Caption         =   "Depósitos Cierre O/F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2730
         TabIndex        =   208
         Top             =   5775
         Width           =   5010
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
            Left            =   1365
            TabIndex        =   210
            Top             =   570
            Width           =   175
         End
         Begin VB.CommandButton Command5 
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
            Left            =   1365
            TabIndex        =   209
            Top             =   255
            Width           =   175
         End
         Begin VB.Label Label103 
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
            Height          =   285
            Left            =   1050
            TabIndex        =   216
            Top             =   570
            Width           =   315
         End
         Begin VB.Label Label102 
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
            Left            =   1680
            TabIndex        =   215
            Top             =   570
            Width           =   3210
         End
         Begin VB.Label Label101 
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
            Height          =   285
            Left            =   1050
            TabIndex        =   214
            Top             =   255
            Width           =   315
         End
         Begin VB.Label Label108 
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
            Left            =   1680
            TabIndex        =   213
            Top             =   255
            Width           =   3210
         End
         Begin VB.Label Label92 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada: "
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
            Left            =   -300
            TabIndex        =   212
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label Label91 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Consumo: "
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
            Left            =   -300
            TabIndex        =   211
            Top             =   615
            Width           =   1320
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
         Height          =   300
         Left            =   5850
         TabIndex        =   8
         Text            =   " "
         Top             =   250
         Width           =   1875
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   6720
         TabIndex        =   44
         Text            =   " "
         Top             =   5365
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   6720
         TabIndex        =   43
         Text            =   " "
         Top             =   4995
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   6720
         TabIndex        =   42
         Text            =   " "
         Top             =   4635
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   6720
         TabIndex        =   41
         Text            =   " "
         Top             =   4275
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   6720
         TabIndex        =   40
         Text            =   " "
         Top             =   3915
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   6720
         TabIndex        =   39
         Text            =   " "
         Top             =   3555
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   6720
         TabIndex        =   38
         Text            =   " "
         Top             =   3195
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6720
         TabIndex        =   37
         Text            =   " "
         Top             =   2835
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   6720
         TabIndex        =   36
         Text            =   " "
         Top             =   2475
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6720
         TabIndex        =   35
         Text            =   " "
         Top             =   2115
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   6720
         TabIndex        =   34
         Text            =   " "
         Top             =   1755
         Width           =   945
      End
      Begin VB.TextBox TextC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   6720
         TabIndex        =   33
         Text            =   " "
         Top             =   1395
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   5040
         TabIndex        =   32
         Text            =   " "
         Top             =   5365
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   5040
         TabIndex        =   31
         Text            =   " "
         Top             =   4995
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   5040
         TabIndex        =   30
         Text            =   " "
         Top             =   4635
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   5040
         TabIndex        =   29
         Text            =   " "
         Top             =   4275
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   5040
         TabIndex        =   28
         Text            =   " "
         Top             =   3915
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   5040
         TabIndex        =   27
         Text            =   " "
         Top             =   3555
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   5040
         TabIndex        =   26
         Text            =   " "
         Top             =   3195
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5040
         TabIndex        =   25
         Text            =   " "
         Top             =   2835
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   5040
         TabIndex        =   24
         Text            =   " "
         Top             =   2475
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5040
         TabIndex        =   23
         Text            =   " "
         Top             =   2115
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5040
         TabIndex        =   22
         Text            =   " "
         Top             =   1755
         Width           =   945
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   5040
         TabIndex        =   21
         Text            =   " "
         Top             =   1395
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   3360
         TabIndex        =   20
         Text            =   " "
         Top             =   5365
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   3360
         TabIndex        =   19
         Text            =   " "
         Top             =   4995
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   3360
         TabIndex        =   18
         Text            =   " "
         Top             =   4635
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   3360
         TabIndex        =   17
         Text            =   " "
         Top             =   4275
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   3360
         TabIndex        =   16
         Text            =   " "
         Top             =   3915
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   15
         Text            =   " "
         Top             =   3555
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   3360
         TabIndex        =   14
         Text            =   " "
         Top             =   3195
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   3360
         TabIndex        =   13
         Text            =   " "
         Top             =   2835
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   3360
         TabIndex        =   12
         Text            =   " "
         Top             =   2475
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3360
         TabIndex        =   11
         Text            =   " "
         Top             =   2115
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Text            =   " "
         Top             =   1755
         Width           =   945
      End
      Begin VB.TextBox TextA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3360
         TabIndex        =   9
         Top             =   1395
         Width           =   945
      End
      Begin VB.TextBox Text7 
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
         Left            =   3300
         TabIndex        =   7
         Text            =   " "
         Top             =   250
         Width           =   1875
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
         Height          =   300
         Left            =   800
         TabIndex        =   6
         Text            =   " "
         Top             =   250
         Width           =   1875
      End
      Begin VB.Label Label107 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   75
         TabIndex        =   232
         Top             =   6435
         Width           =   645
      End
      Begin VB.Label Label83 
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
         Left            =   45
         TabIndex        =   231
         Top             =   6660
         Width           =   2535
      End
      Begin VB.Label Label106 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6070
         TabIndex        =   230
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label105 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mínimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6085
         TabIndex        =   229
         Top             =   950
         Width           =   620
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6725
         TabIndex        =   228
         Top             =   1135
         Width           =   945
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   6725
         TabIndex        =   227
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label99 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4400
         TabIndex        =   226
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mínimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4410
         TabIndex        =   225
         Top             =   950
         Width           =   620
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   5045
         TabIndex        =   224
         Top             =   1135
         Width           =   945
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5045
         TabIndex        =   223
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label95 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   222
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label94 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mínimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2750
         TabIndex        =   221
         Top             =   950
         Width           =   620
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3365
         TabIndex        =   220
         Top             =   1135
         Width           =   945
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3365
         TabIndex        =   219
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label84 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A.Int.:"
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
         Left            =   4820
         TabIndex        =   198
         Top             =   300
         Width           =   1020
      End
      Begin VB.Line Line4 
         X1              =   2640
         X2              =   7800
         Y1              =   5715
         Y2              =   5715
      End
      Begin VB.Label Label81 
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
         Left            =   6120
         TabIndex        =   193
         Top             =   5430
         Width           =   540
      End
      Begin VB.Label Label80 
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
         Left            =   6120
         TabIndex        =   192
         Top             =   5070
         Width           =   540
      End
      Begin VB.Label Label79 
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
         Left            =   6120
         TabIndex        =   191
         Top             =   4710
         Width           =   540
      End
      Begin VB.Label Label78 
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
         Left            =   6120
         TabIndex        =   190
         Top             =   4350
         Width           =   540
      End
      Begin VB.Label Label77 
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
         Left            =   6120
         TabIndex        =   189
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label Label76 
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
         Left            =   6120
         TabIndex        =   188
         Top             =   3630
         Width           =   540
      End
      Begin VB.Label Label75 
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
         Left            =   6120
         TabIndex        =   187
         Top             =   3270
         Width           =   540
      End
      Begin VB.Label Label74 
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
         Left            =   6120
         TabIndex        =   186
         Top             =   2910
         Width           =   540
      End
      Begin VB.Label Label73 
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
         Left            =   6120
         TabIndex        =   185
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label Label72 
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
         Left            =   6120
         TabIndex        =   184
         Top             =   2190
         Width           =   540
      End
      Begin VB.Label Label71 
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
         Left            =   6120
         TabIndex        =   183
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label70 
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
         Height          =   300
         Left            =   6120
         TabIndex        =   182
         Top             =   1470
         Width           =   540
      End
      Begin VB.Label Label69 
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
         Left            =   4440
         TabIndex        =   181
         Top             =   5430
         Width           =   540
      End
      Begin VB.Label Label68 
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
         Left            =   4440
         TabIndex        =   180
         Top             =   5070
         Width           =   540
      End
      Begin VB.Label Label67 
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
         Left            =   4440
         TabIndex        =   179
         Top             =   4710
         Width           =   540
      End
      Begin VB.Label Label66 
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
         Left            =   4440
         TabIndex        =   178
         Top             =   4350
         Width           =   540
      End
      Begin VB.Label Label65 
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
         Left            =   4440
         TabIndex        =   177
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label Label64 
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
         Left            =   4440
         TabIndex        =   176
         Top             =   3630
         Width           =   540
      End
      Begin VB.Label Label63 
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
         Left            =   4440
         TabIndex        =   175
         Top             =   3270
         Width           =   540
      End
      Begin VB.Label Label62 
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
         Left            =   4440
         TabIndex        =   174
         Top             =   2910
         Width           =   540
      End
      Begin VB.Label Label61 
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
         Left            =   4440
         TabIndex        =   173
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label Label60 
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
         Left            =   4440
         TabIndex        =   172
         Top             =   2190
         Width           =   540
      End
      Begin VB.Label Label59 
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
         Left            =   4440
         TabIndex        =   171
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label58 
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
         Height          =   300
         Left            =   4440
         TabIndex        =   170
         Top             =   1470
         Width           =   540
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
         Left            =   2760
         TabIndex        =   169
         Top             =   5430
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
         Left            =   2760
         TabIndex        =   168
         Top             =   5070
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
         Left            =   2760
         TabIndex        =   167
         Top             =   4710
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
         Left            =   2760
         TabIndex        =   166
         Top             =   4350
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
         Left            =   2760
         TabIndex        =   165
         Top             =   3990
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
         Left            =   2760
         TabIndex        =   164
         Top             =   3630
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
         Left            =   2760
         TabIndex        =   163
         Top             =   3270
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
         Left            =   2760
         TabIndex        =   162
         Top             =   2910
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
         Left            =   2760
         TabIndex        =   161
         Top             =   2550
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
         Left            =   2760
         TabIndex        =   160
         Top             =   2190
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
         Left            =   2760
         TabIndex        =   159
         Top             =   1830
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
         Height          =   300
         Left            =   2760
         TabIndex        =   158
         Top             =   1470
         Width           =   540
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   2650
         X2              =   2650
         Y1              =   660
         Y2              =   7000
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº Control   Nº Control   Mediciones A    Mediciones B    Mediciones C"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Top             =   630
         Width           =   7785
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1320
         Y1              =   840
         Y2              =   6400
      End
      Begin VB.Label Lis2 
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
         Index           =   29
         Left            =   2160
         TabIndex        =   157
         Top             =   6120
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   29
         Left            =   1680
         TabIndex        =   156
         Top             =   6120
         Width           =   360
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Left            =   1320
         TabIndex        =   155
         Top             =   6150
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   28
         Left            =   2160
         TabIndex        =   154
         Top             =   5760
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   28
         Left            =   1680
         TabIndex        =   153
         Top             =   5760
         Width           =   360
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "29"
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
         Left            =   1320
         TabIndex        =   152
         Top             =   5790
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   27
         Left            =   2160
         TabIndex        =   151
         Top             =   5400
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   27
         Left            =   1680
         TabIndex        =   150
         Top             =   5400
         Width           =   360
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "28"
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
         Left            =   1320
         TabIndex        =   149
         Top             =   5430
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   26
         Left            =   2160
         TabIndex        =   148
         Top             =   5040
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   26
         Left            =   1680
         TabIndex        =   147
         Top             =   5040
         Width           =   360
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "27"
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
         Left            =   1320
         TabIndex        =   146
         Top             =   5070
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   25
         Left            =   2160
         TabIndex        =   145
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   25
         Left            =   1680
         TabIndex        =   144
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "26"
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
         Left            =   1320
         TabIndex        =   143
         Top             =   4710
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   24
         Left            =   2160
         TabIndex        =   142
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   24
         Left            =   1680
         TabIndex        =   141
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "25"
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
         Left            =   1320
         TabIndex        =   140
         Top             =   4350
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   23
         Left            =   2160
         TabIndex        =   139
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   23
         Left            =   1680
         TabIndex        =   138
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "24"
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
         Left            =   1320
         TabIndex        =   137
         Top             =   3990
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   22
         Left            =   2160
         TabIndex        =   136
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   22
         Left            =   1680
         TabIndex        =   135
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "23"
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
         Left            =   1320
         TabIndex        =   134
         Top             =   3630
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   21
         Left            =   2160
         TabIndex        =   133
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   21
         Left            =   1680
         TabIndex        =   132
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "22"
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
         Left            =   1320
         TabIndex        =   131
         Top             =   3270
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   20
         Left            =   2160
         TabIndex        =   130
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   20
         Left            =   1680
         TabIndex        =   129
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "21"
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
         Left            =   1320
         TabIndex        =   128
         Top             =   2910
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   19
         Left            =   2160
         TabIndex        =   127
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   19
         Left            =   1680
         TabIndex        =   126
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label44 
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
         Left            =   1320
         TabIndex        =   125
         Top             =   2550
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   18
         Left            =   2160
         TabIndex        =   124
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   18
         Left            =   1680
         TabIndex        =   123
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Label43 
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
         Left            =   1320
         TabIndex        =   122
         Top             =   2190
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   17
         Left            =   2160
         TabIndex        =   121
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   17
         Left            =   1680
         TabIndex        =   120
         Top             =   1800
         Width           =   360
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
         Left            =   1320
         TabIndex        =   119
         Top             =   1830
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   16
         Left            =   2160
         TabIndex        =   118
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   16
         Left            =   1680
         TabIndex        =   117
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label41 
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
         Left            =   1320
         TabIndex        =   116
         Top             =   1470
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   15
         Left            =   2160
         TabIndex        =   115
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   15
         Left            =   1680
         TabIndex        =   114
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label40 
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
         Left            =   1320
         TabIndex        =   113
         Top             =   1110
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   14
         Left            =   840
         TabIndex        =   112
         Top             =   6120
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   14
         Left            =   360
         TabIndex        =   111
         Top             =   6120
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
         Left            =   0
         TabIndex        =   110
         Top             =   6150
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   13
         Left            =   840
         TabIndex        =   109
         Top             =   5760
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   13
         Left            =   360
         TabIndex        =   108
         Top             =   5760
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
         Left            =   0
         TabIndex        =   107
         Top             =   5790
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   12
         Left            =   840
         TabIndex        =   106
         Top             =   5400
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   12
         Left            =   360
         TabIndex        =   105
         Top             =   5400
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
         Left            =   0
         TabIndex        =   104
         Top             =   5430
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   11
         Left            =   840
         TabIndex        =   103
         Top             =   5040
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   11
         Left            =   360
         TabIndex        =   102
         Top             =   5040
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
         Left            =   0
         TabIndex        =   101
         Top             =   5070
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   10
         Left            =   840
         TabIndex        =   100
         Top             =   4680
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   10
         Left            =   360
         TabIndex        =   99
         Top             =   4680
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
         Left            =   0
         TabIndex        =   98
         Top             =   4710
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   9
         Left            =   840
         TabIndex        =   97
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   9
         Left            =   360
         TabIndex        =   96
         Top             =   4320
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
         Left            =   0
         TabIndex        =   95
         Top             =   4350
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   8
         Left            =   840
         TabIndex        =   94
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   8
         Left            =   360
         TabIndex        =   93
         Top             =   3960
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
         Left            =   0
         TabIndex        =   92
         Top             =   3990
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   7
         Left            =   840
         TabIndex        =   91
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   7
         Left            =   360
         TabIndex        =   90
         Top             =   3600
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
         Left            =   0
         TabIndex        =   89
         Top             =   3630
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   6
         Left            =   840
         TabIndex        =   88
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   6
         Left            =   360
         TabIndex        =   87
         Top             =   3240
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
         Left            =   0
         TabIndex        =   86
         Top             =   3270
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   5
         Left            =   840
         TabIndex        =   85
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   5
         Left            =   360
         TabIndex        =   84
         Top             =   2880
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
         Left            =   0
         TabIndex        =   83
         Top             =   2910
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   4
         Left            =   840
         TabIndex        =   82
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   4
         Left            =   360
         TabIndex        =   81
         Top             =   2520
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
         Left            =   0
         TabIndex        =   80
         Top             =   2550
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   3
         Left            =   840
         TabIndex        =   79
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   3
         Left            =   360
         TabIndex        =   78
         Top             =   2160
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
         Left            =   0
         TabIndex        =   77
         Top             =   2190
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   2
         Left            =   840
         TabIndex        =   76
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   2
         Left            =   360
         TabIndex        =   75
         Top             =   1800
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
         Left            =   0
         TabIndex        =   74
         Top             =   1830
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   1
         Left            =   840
         TabIndex        =   73
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   1
         Left            =   360
         TabIndex        =   72
         Top             =   1440
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
         Left            =   0
         TabIndex        =   71
         Top             =   1470
         Width           =   300
      End
      Begin VB.Label Lis2 
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
         Index           =   0
         Left            =   840
         TabIndex        =   70
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Lis1 
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
         Index           =   0
         Left            =   360
         TabIndex        =   69
         Top             =   1095
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
         Left            =   0
         TabIndex        =   68
         Top             =   1110
         Width           =   300
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A.Ext.:"
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
         Left            =   2280
         TabIndex        =   66
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label24 
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
         Height          =   300
         Left            =   0
         TabIndex        =   65
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Encabezado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   120
      TabIndex        =   52
      Top             =   960
      Width           =   3615
      Begin VB.TextBox Text8 
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
         Left            =   315
         TabIndex        =   201
         Text            =   " "
         Top             =   2865
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
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
         TabIndex        =   4
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
         Height          =   1650
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3600
         Width           =   3330
      End
      Begin VB.CommandButton Command7 
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
         Left            =   2640
         TabIndex        =   1
         Top             =   1125
         Width           =   165
      End
      Begin VB.TextBox Text2 
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
         Left            =   360
         TabIndex        =   2
         Text            =   " "
         Top             =   600
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
         Height          =   285
         Left            =   1290
         TabIndex        =   3
         Top             =   600
         Width           =   165
      End
      Begin VB.TextBox Text3 
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
         Left            =   1125
         TabIndex        =   0
         Text            =   " "
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label97 
         BackStyle       =   0  'Transparent
         Caption         =   "BR:"
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
         Left            =   2730
         TabIndex        =   242
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label93 
         Alignment       =   1  'Right Justify
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
         Left            =   2670
         TabIndex        =   241
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label90 
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
         Height          =   285
         Left            =   3090
         TabIndex        =   207
         ToolTipText     =   "Número de Revisión"
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0C0&
         Caption         =   "Cantidad Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   206
         Top             =   3300
         Width           =   1065
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
         Left            =   360
         TabIndex        =   61
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label86 
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
         Height          =   405
         Left            =   2850
         TabIndex        =   200
         Top             =   1200
         Width           =   200
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         TabIndex        =   199
         ToolTipText     =   "Número de Revisión"
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dictámen"
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
         Left            =   1680
         TabIndex        =   63
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tensión"
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
         TabIndex        =   62
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         TabIndex        =   60
         Top             =   3300
         Width           =   1410
      End
      Begin VB.Label Label13 
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
         Height          =   525
         Left            =   150
         TabIndex        =   59
         Top             =   2040
         Width           =   3345
      End
      Begin VB.Label Label7 
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
         Left            =   150
         TabIndex        =   58
         Top             =   1680
         Width           =   2480
      End
      Begin VB.Label Label5 
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
         Left            =   0
         TabIndex        =   57
         Top             =   1440
         Width           =   780
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
         Left            =   -120
         TabIndex        =   56
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Serie"
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
         Left            =   80
         TabIndex        =   55
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/F"
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
         Left            =   1680
         TabIndex        =   54
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Label15 
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
         Left            =   2040
         TabIndex        =   53
         ToolTipText     =   "Doble clic Para Ver Situación de Reservas Pendientes Transferidas"
         Top             =   600
         Width           =   1440
      End
   End
   Begin VB.Label Label88 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dictámen"
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
      Left            =   2295
      TabIndex        =   205
      Top             =   4050
      Width           =   1065
   End
   Begin VB.Label Label87 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "fgdfgdfg"
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
      Left            =   2055
      TabIndex        =   204
      Top             =   4110
      Width           =   1065
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5040
      Picture         =   "REGICON07.frx":3740
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "REGICON07.frx":3A4A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Registro de Control Final"
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
      TabIndex        =   51
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "REGICON07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUMECONTROL&
Dim TXTRAZA_REGICON07$



Function VALIDA_MEDICION%()
  '
  'SI ESTA TODO CORRECTO DEVUELVE 1 SI SE PRODUCE ERROR DEVUELVE 0
  'SE VALIDA QUE HAYA VALOR MAXIMO YA QUE EL MINIMO VA  A DEVOLVER 0
  '
  VALIDA_MEDICION% = 0
  MINIA = XVALO(Label82(0))
  MAXIA = XVALO(Label82(1))
  If MAXIA > 0.005 Then CONTROLA_A% = 1 'PARA QUE SI NO TIENE VALOR MAXIMO NO CONTROLE
  '
  MINIB = XVALO(Label96(0))
  MAXIB = XVALO(Label96(1))
  If MAXIB > 0.005 Then CONTROLA_B% = 1 'PARA QUE SI NO TIENE VALOR MAXIMO NO CONTROLE
  '
  MINIC = XVALO(Label100(0))
  MAXIC = XVALO(Label100(1))
  If MAXIC > 0.005 Then CONTROLA_C% = 1 'PARA QUE SI NO TIENE VALOR MAXIMO NO CONTROLE
  '
  'SI NO HAY INGRESOS DE RANGO DE VALIDACIONES NO CONTROLA
  If (CONTROLA_A% + CONTROLA_B% + CONTROLA_C%) < 1 Then Exit Function
  '
  For i& = 0 To TextA.Count - 1
    If CONTROLA_A% > 0 Then
      MEDAX = XVALO(TextA(i&))
      If Abs(MEDAX) > 0.005 Then
        If MEDAX < MINIA Or MEDAX > MAXIA Then
           POS$ = "00" & SAFETEXT$(i& + 1): POS$ = Right(POS$, 2)
           Call MENSERR(24, "Valor Ingresado: " & MEDAX & " En Medición 'A' En CH" & POS$ & "\Fuera de Limites Permitidos")
           VALIDA_MEDICION% = 1
           On Error Resume Next
           TextA(i&).SetFocus
           On Error GoTo 0
           Exit Function
        End If
      End If
    End If
    '
    If CONTROLA_B% > 0 Then
      MEDBX = XVALO(TextB(i&))
      If Abs(MEDBX) > 0.005 Then
        If MEDBX < MINIB Or MEDBX > MAXIB Then
           POS$ = "00" & SAFETEXT$(i& + 1): POS$ = Right(POS$, 2)
           Call MENSERR(24, "Valor Ingresado: " & MEDBX & " En Medición 'B' En CH" & POS$ & "\Fuera de Limites Permitidos")
           VALIDA_MEDICION% = 1
           On Error Resume Next
           TextB(i&).SetFocus
           On Error GoTo 0
           Exit Function
        End If
      End If
    End If
    '
    If CONTROLA_C% > 0 Then
      MEDCX = XVALO(TextC(i&))
      If Abs(MEDCX) > 0.005 Then
        If MEDCX < MINIC Or MEDCX > MAXIC Then
           POS$ = "00" & SAFETEXT$(i& + 1): POS$ = Right(POS$, 2)
           Call MENSERR(24, "Valor Ingresado: " & MEDCX & " En Medición 'C' En CH" & POS$ & "\Fuera de Limites Permitidos")
           VALIDA_MEDICION% = 1
           On Error Resume Next
           TextC(i&).SetFocus
           On Error GoTo 0
           Exit Function
        End If
      End If
   End If
        
  Next i&

End Function

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command11_Click(Index As Integer)
   Label109(Index) = ""
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    YATRANSFIRIO% = 0
    '
    ' VALIDACION POR FECHA
    VDEVV$ = TRIM$(Text2.TEXT)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    ' VALIDACION DE NUMERO DE SERIE
    NUSERIE$ = TRIM$(Text3)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    ' VALIDACION DE DICTAMEN
    If UCase$(Left$(Text5, 4)) <> "APRO" Then
      If Left$(Text5, 4) <> "RECH" Then
        Call MENSERR(24, "Falta Informar Dictamen o No Válido")
        GoTo 99
      End If
    End If
    '
    ' Validacion de Tecnico
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Falta Informar Técnico")
      GoTo 99
    End If
    '
    ' Validacion de Tension
    If TRIM$(Label8) = "" Then
      Call MENSERR(24, "Falta Tensión del Equipo")
      GoTo 99
    End If
    '
    'VALIDA SI SE HAN INGRESADO MOVIMIENTOS
    VACIO% = 1
    For i& = 0 To 3
       If TRIM$(Label109(i&)) <> "" Then VACIO% = 0
    Next i&
    If VACIO% > 0 Then
      OPX% = COMALTER%("No se Han Ingresado Instrumentos.\\Cancelar Grabación\Continuar La Grabación")
      If OPX% < 2 Then GoTo 99
    End If
    '
    'VALIDO LOS DATOS INGRESADOS EN LA MEDICION
    If VALIDA_MEDICION% > 0 Then
      GoTo 99
    End If
    '
    DEPOIMPORTADO% = 0
    
    Dim obj As New Art_Importados
    CODI$ = Label7.Caption
    Call obj.VeriImportado(CODI$) 'VERIFICA QUE NO SEA IMPORTADO

    If XVALO(Label93) > 0 And obj.EsImportado > 0 Then 'IMPORTADO
       DEPOIMPORTADO% = XVALO(CONTROL("", "DEPOIMPO"))
       If DEPOIMPORTADO% < 1 Then
          Call COMUNI("Falta Definir Depósito de Cierre Para Productos Ingresados por Recepción")
          GoTo 99
       End If
    End If
    
    ' AQUI COMIENZAN VALIDACIONES POR CIERRE DE ORDEN DE FABRICACION
    CIERRAOF% = 0
    If UCase(Left$(Text5.TEXT, 4)) <> "APRO" Then GoTo 25  ' NO ESTA APROBADO
      DEPOANTER% = XVALO(VALOBADA("TANUMSE", "DEPONUMSE", "NumeroSerie = '" & NUSERIE$ & "'", "NOMESS"))
    If DEPOANTER% >= 0 Then GoTo 25                 ' YA ESTABA EN DEPOSITO REAL - PRE-APROBADO
      CONDI$ = "CodiMovi = 'OF' "
      CONDI$ = CONDI$ + " AND IdenLote = '" & NUSERIE$ & "' "
      CONDI$ = CONDI$ + " AND DoPriCor = '" & Label15 & "' "
      CONDI$ = CONDI$ + " AND CantIngre > 0 "
      ENTRADO% = CantRegistros("MOVISTO", CONDI$)
    If ENTRADO% > 0 Then GoTo 25
    '
    ' NO REGISTRADO INGRESO DE EQUIPO - COMIENZA VALIDACION PARA EL CIERRE DE LA o/f
    NORFAX$ = Label15
    'COMENTADO PARA QUE CONSUMA DEL 54 COMO APARECE POR DEFECTO QUE ES DE DONDE DEBE CONSUMIR
''    If ValidaArmador%(NORFAX$) > 0 Then
''       DEPARMA% = XVALO(CONTROL("", "DEPOARMA"))
''       If XVALO(Label103) <> DEPARMA% Then
''         Call MENSERR(24, "No Válido Cambiar de Depósito\Orden de Fabricación Asignado a Armador Externo\Depósito de Consumo: " & DEPARMA%)
''         Label103 = TRIM$(Str$(DEPARMA%))
''         GoTo 99
''       End If
''    End If
    DEPOTALLER% = XVALO(Label103)
    '
    DEPOENTRA% = XVALO(Label101) ' SI ES IMPORTADO ASIGNA EL DEPOSITO DE ENTRADA EL CONFIGURADO PARA ESTOS
    If DEPOIMPORTADO% > 0 Then DEPOENTRA% = DEPOIMPORTADO%
    
    If DEPOENTRA% < 1 Then
       Call MENSERR(24, "Depósito de Entrada no Válido")
       GoTo 99
    End If
    '
    Screen.MousePointer = 11
    CONDI$ = "IdSubOr = '" & NORFAX$ & "'"
    CIXI% = VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS")
    CODICONJ$ = CODEXT$(CIXI%)
    VACONTROL% = 1
    Call CARCOMPO(CODICONJ$, NORFAX$, 1, VACONTROL%, 5)
    If VACONTROL% < 0 Then               ' SI CANCELA SE VA
       OPXX% = COMALTER%("Cancelado Cierre de O/F - Opciones de Trabajo:\\Cancelar Reg-Con-Fin\Dictamen sin Ingreso de Equipo")
       If OPXX% <> 2 Then GoTo 99 ' CANCELADO
    End If
    '
    OPXX% = 3: CONTINUA% = 1
    If CONTROL("VANEDECE", "DEPOCEN") = "SI" Then ' BLOQUEO DE NEGATIVOS ON
        'ARMO CONSULTA PARA OBTENER LOS STOCKS
        INGRECONDI% = 0
        For i& = 1 To ULTREG("!MATEROF")          ' QUE
              XXX$ = REGLEIDO$("!MATEROF", i&)
              CI% = CVI(Mid$(XXX$, 1, 2))
              If i& = 1 Then
                INGRECONDI% = 1
                CONDI$ = " (Cod_Inte = " & CI%
              Else
                CONDI$ = CONDI$ + " or Cod_Inte = " & CI%
              End If
        Next i&
        '
        Dim STOCKSALI(32767)
        Dim STOCKINGRE(32767)
        '
        If INGRECONDI% = 1 Then
           CONDI$ = CONDI$ + " )"
           'CARGA EL STOCK INGRESADO EN UN VECTOR PARA LOS CODIGOS COMPONENTES
           'SE HACE POR SEPARADO POR QUE DIFIERE LA CONDICION
           SQLX$ = " Select COD_INTE, sum(CantIngre) AS SUINGRE "
           SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) "
           SQLX$ = SQLX$ + " Where " & CONDI$
           SQLX$ = SQLX$ + " AND DOSECCOR = '" & TRIM$(NORFAX$) & "' And DepoMovi = " & DEPOTALLER%
           SQLX$ = SQLX$ + "  GROUP BY COD_INTE "
           Set MIRS = READSET(SQLX$)
           With MIRS
              Do While Not .EOF
                CIXI% = XVALO(!COD_INTE)
                If CIXI% > 0 Then STOCKINGRE(CIXI%) = XVALO(!SUINGRE)
                .MoveNext
              Loop
           End With
           MIRS.Close
           Set MIRS = Nothing
           'CARGA EL STOCK SALIDO EN UN VECTOR PARA LOS CODIGOS COMPONENTES
           'SE HACE POR SEPARADO POR QUE DIFIERE LA CONDICION
           SQLX$ = " Select COD_INTE, SUM(CantSalid) as SUSALI "
           SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) "
           SQLX$ = SQLX$ + " Where " & CONDI$
           SQLX$ = SQLX$ + " AND DOPRICOR = '" & TRIM$(NORFAX$) & "' And DepoMovi = " & DEPOTALLER%
           SQLX$ = SQLX$ + "  GROUP BY COD_INTE "
           Set MIRS = READSET(SQLX$)
           With MIRS
              Do While Not .EOF
                CIXI% = XVALO(!COD_INTE)
                If CIXI% > 0 Then STOCKSALI(CIXI%) = XVALO(!SUSALI)
                .MoveNext
              Loop
           End With
           MIRS.Close
           Set MIRS = Nothing
        End If
        '
        For i& = 1 To ULTREG("!MATEROF")          ' QUE
          XXX$ = REGLEIDO$("!MATEROF", i&)
          CI% = CVI(Mid$(XXX$, 1, 2))
          UNIME$ = Mid$(XXX$, 3, 2)
          USOUNIT# = CVD(Mid$(XXX$, 9, 8))
          NECESI# = USOUNIT#
          '
          ' stock transferido a deposito taller (54)
'5
'          CONDI$ = "Cod_Inte = " & CI% & " AND DOSECCOR = '" & TRIM$(NORFAX$) & "' And DepoMovi = " & DEPOTALLER%
'          STOXXX1# = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
'          '
'          ' stock consumido de deposito taller (54)
'          CONDI$ = "CodiMovi = 'CF' AND Cod_Inte = " & CI% & " AND DOPRICOR = '" & TRIM$(NORFAX$) & "' AND DepoMovi = " & DEPOTALLER%
'          STOXXX2# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          '
5         STOXXX1# = STOCKINGRE(CI%)
          STOXXX2# = STOCKSALI(CI%)
          
          ' saldo en taller
          SOTOCKK# = STOXXX1# - STOXXX2#
          If (NECESI# - SOTOCKK#) > 0.05 Then
             
             If OPXX% > 2 Then
                'SI ENTRO ANTES DE PRESENTAR MENSAJE SE FIJA SI TIENE MATERIAL EN EL 32 Y TRANFIERE AL 54
                If YATRANSFIRIO% < 1 Then
                  'SI LA CANTIDAD DEL STOCK(54) ES NEGATIVA TRATA DE TRANSFERIR PARA COMPLETAR EN POSITIVO
                  '+ LA CANTIDAD DE USO UNITARIO (EJ. SI EL STOCK(54) ES - 5 Y EL USO ES 1
                  'LA NECESIDAD A TRANSFERIR DEL 32 AL 54 VA A SER DE 6
                  NECESIDAD = USOUNIT#
                  If SOTOCKK# < FRACMIN(CI%) Then NECESIDAD = USOUNIT# + Abs(SOTOCKK#)
                  Call TRANSFIERE_COMPONENTES(CI%, CIXI%, NUSERIE$, NECESIDAD, Resultado%)
                  YATRANSFIRIO% = 1
                  If Resultado% = -1 Then
                    TXT$ = "Imposible Continuar. El Stock Del Componente '" & CODEXT$(CI%) & "'\No Es Suficiente En El Depósito " & DEPOTALLER%
                    Call COMUNI(TXT$)
                    CONTINUA% = 0
                  End If
                End If
                '
'                TTXX$ = "Alerta al Cerrar Esta O.F.: " & NORFAX$
'                TTXX$ = TTXX$ + "\Código: " & CODEXT$(CI%) & " con Faltantes en Depósito " & DEPOTALLER%
'                TTXX$ = TTXX$ + "\Opciones de Trabajo:\\Cancelar Reg-Con-Fin\Dictamen sin Ingreso de Equipo\Ver Detalle de Faltantes"
'90                OPXX% = COMALTER%(TTXX$)
'                If OPXX% < 2 Then GoTo 99 ' CANCELADO
'                If OPXX% = 3 Then
'                    Call INFOCONSU(NORFAX$, 1)
'                    GoTo 90
'                End If
             End If
          End If
100       YATRANSFIRIO% = 0 'SOLUCION PROVISORIA PARA QUE CONTINUE SIN PRESENTAR MENSAJE
        Next i&
    End If
    '
    If CONTINUA% = 0 Then GoTo 98
    '
    If OPXX% = 3 Then CIERRAOF% = 1 ' SI LLEGA CON OPXX%=3 CIERRA
    '
    ' FIN VALIDACIONES PARA EL CIERRE DE LA O/F
    '
    If CIERRAOF% = 1 Then
       ' 1.- MOVIMIENTO DE INGRESO DE EQUIPO
          Call ENTRAEQUIPO
       ' 2.- GRABAR DEPOSITO DE INGRESO EN TANUMSE
          SQLX$ = "NumeroSerie = '" & NUSERIE$ & "'"
10:       Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", SQLX$, DEPOENTRA%, REGAF&)
          '**** ESTO ES UNA SOLUCION PROVISORIA !!!!
          'SE INCORPORA PARA CAPTURAR QUE NO SE ACTUALIZA EL DEPOSITO Y POSIBILITAR REINTENTARLO
          DEPOVERIF% = XVALO(VALOBADA("TANUMSE", "DEPONUMSE", SQLX$))
          If DEPOVERIF% <> DEPOENTRA% Then
            OP% = COMALTER%("No se Ha Podido Ejecutar La Consulta SQL.\¿Desea Intentarlo Nuevamente?\\Si, Reintentar\No, Continuar")
            If OP% = 1 Then GoTo 10
          End If
          '**** FIN SOLUCION PROVISORIA
       ' 3.- ANOTA EN EL ANOTADOR DE LA ORDEN DE FABRICACION
          Call ANOTAENORFA
       ' 4.- ACTUALIZAR SALDO DE LA ORDEN DE FABRICACION
          CONDI$ = "CODIMOVI = 'OF' AND DOPRILAR = '" & NORFAX$ & "'"
          TOFABRI = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
          TOFABRI = TOFABRI - XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
          
          If TOFABRI < 0 Then TOFABRI = 0
          Call ACTUREGISTRO("ORDEFABR", "CanApro", "IdSubOr = '" & NORFAX$ & "'", TOFABRI, REGAFE&, "NOMESS")
          
       ' 5.- MOVIMIENTO DE CONSUMO DE COMPONENTES CONSOLIDADO
'          Call CONSUCOMPO
          Call CONSUCOMPO_1
          
       ' 6.- CIERRE DE LA ORDEN DE FABRICACION
          CANPROYE& = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFAX$ & "'"))
          CANREALE& = XVALO(VALOBADA("ORDEFABR", "CanApro", "IdSubOr = '" & NORFAX$ & "'"))
          '
          If CANREALE& >= CANPROYE& Then
              CONDI$ = "IdSubOr = '" & NORFAX$ & "'"
              HOII% = HOY(HOS$)
              FECHAHOY = FETEXCOR1$(HOII%)
             
              Call ACTUREGISTRO("ORDEFABR", "StatuOrf", CONDI$, 3, REGAF&, "NOMESS")
              Call ACTUREGISTRO("ORDEFABR", "FechCierr", CONDI$, FECHAHOY, REGAF&, "NOMESS")
             
              ' Cancela Operaciones Programadas
              Call BORRAREGISTROS("OperFab", CONDI$, RAFE&, "NOMESS")
              
              ' Cancela Reservas de Componentes
              Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
              
              ' Cancela Orden en Indice de Ordenes Pendientes
              NORFAX$ = AJUSTI$(NORFAX$, 12)
              Call FILTERDELETE("ECORDEP", 1, NORFAX$)
              Call FRESHECHO("ECORDEP")
          End If
          Call CIERRARCH("*.*")
          Call LIBARCH("ECORDEP")
          Call CIERRARCH("*.*")
          '
    End If
    '
25  Screen.MousePointer = 11
    '
'    SQLX$ = "NumeroSerie = '" & NUSERIE$ & "'"
'    Call ACTUREGISTRO("TANUMSE", "Ten_Ensa", SQLX$, XVALO(Text8), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "Dic_Ensa", SQLX$, UCase$(Left$(TRIM$(Text5), 4)), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "Tec_Ensa", SQLX$, Left$(TRIM$(Text6), 32), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "Aext_Ensa", SQLX$, Left$(TRIM$(Text7), 32), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "Obse_Ensa", SQLX$, Text4, REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "Fech_Ensa", SQLX$, TRIM$(Text2), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "AINT_Ensa", SQLX$, Left$(TRIM$(Text1), 32), REGAF&)
'
'    Call ACTUREGISTRO("TANUMSE", "INSTRUMENTO1", SQLX$, TRIM$(Label109(0)), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "INSTRUMENTO2", SQLX$, TRIM$(Label109(1)), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "INSTRUMENTO3", SQLX$, TRIM$(Label109(2)), REGAF&)
'    Call ACTUREGISTRO("TANUMSE", "INSTRUMENTO4", SQLX$, TRIM$(Label109(3)), REGAF&)

    SQLX$ = "SELECT * FROM TANUMSE WHERE NUMEROSERIE='" & NUSERIE$ & "'"
    Dim TANUMSETEMP As ADODB.Recordset
    Set TANUMSETEMP = New ADODB.Recordset
26  On Error Resume Next
    TANUMSETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0
    With TANUMSETEMP
      If Not .EOF Then
        !TEN_ENSA = XVALO(Text8)
        !Dic_Ensa = UCase$(Left$(TRIM$(Text5), 4))
        !Tec_Ensa = Left$(TRIM$(Text6), 32)
        !Aext_Ensa = Left$(TRIM$(Text7), 32)
        !Obse_Ensa = TRIM$(Text4)
        !FECH_ENSA = TRIM$(Text2)
        !AINT_Ensa = Left$(TRIM$(Text1), 32)
        '
        !INSTRUMENTO1 = TRIM$(Label109(0))
        !INSTRUMENTO2 = TRIM$(Label109(1))
        !INSTRUMENTO3 = TRIM$(Label109(2))
        !INSTRUMENTO4 = TRIM$(Label109(3))
        .Update
      End If
    End With
    TANUMSETEMP.Close
    Set TANUMSETEMP = Nothing
    '
    
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Registro de Control Final" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Text6))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text5))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    REVIS% = 0
    SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM CONTROLFIN WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
    SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
    SQLZ$ = SQLZ$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"
    Dim MAXREVTMP As ADODB.Recordset
    'Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
    Set MAXREVTMP = READSET(SQLZ$)
    ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
    MAXREVTMP.Close
    Set MAXREVTMP = Nothing
    '
    'numero de registro de control final
    CONDI$ = "NUMEROSERIE <> ''"
    NUMECONTROL& = 1 + XVALO(VALOBADA("CONTROLFIN", "MAX(NUMCONTROL)", CONDI$, "MOMESS"))
    If NUMECONTROL& < 2 Then
50     NUMECONTROL& = ROUND(XVALO(InputBox$("Ingrese Número", "Inicialización de Numeración")), 0)
       If NUMECONTROL& < 1 Then
          Call COMUNI("Ingreso No Válido")
          GoTo 50
       End If
    End If
    '
    
    SQLX$ = "SELECT REVISION FROM CONTROLFIN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
    SQLX$ = SQLX$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"
    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
    '
    Dim CONTRTMP As ADODB.Recordset
    Set CONTRTMP = READSET(SQLX$)
   
    
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
    
    For U& = 0 To Lis1.Count - 1
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
          !FECH_ENSA = TRIM$(Text2)
          !Revision = REVIS% '\\\OT-08-1064-OD-21/11/08-MEDITEA- GRABA REVISION
          !NUMCONTROL = NUMECONTROL&
          !Tec_Ensa = USERNAME$
          If U& < 1 Then ' graba solo en el primer registro
             !Dic_Ensa = Left$(TRIM$(Text5), 32)  'dictamen de ensayo
             !Obse_Ensa = Text4                   'observaciones
             !Aext_Ensa = Left$(TRIM$(Text7), 32) 'auxiliar externo
             !AINT_Ensa = Left$(TRIM$(Text1), 32) 'auxliar interno
          End If

          .Update
        End With
    Next U&
    CTRLFINTMP.Close
   
    '
    'GRABO LOS DATOS COMO NUEVOS EN MEDIFIN
    SQLX$ = "SELECT * FROM MEDIFIN"
    SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NUSERIE$ & "'"
    Dim MEDIFINTMP As ADODB.Recordset
    Set MEDIFINTMP = New ADODB.Recordset
    MEDIFINTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For i& = 0 To TextA.Count - 1
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
          !MEDB = XVALO(TextB(i&))
          !MEDC = XVALO(TextC(i&))
          '!MARBORRA = 0
          !FECH_ENSA = TRIM$(Text2)
          !Revision = REVIS% '\\\OT-08-1064-OD-21/11/08-MEDITEA- GRABA REVISION
          .Update
        End With
    Next i&
    MEDIFINTMP.Close
    '
    Call COMUNI("Registro Completo. Pulse 'Aceptar' para Imprimir Comprobante.")
    Call PRINTMEDI(1) ' UNO GRABACION  --- 0 IMPRESION SIMPLE
    '
    Command23.Enabled = True
98  Call BLANFORMU
    '
''GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TXTRAZA_REGICON07-" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_REGICON07$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%
99  Screen.MousePointer = 1
    Command2.Enabled = True
    '
End Sub

Sub TRANSFIERE_COMPONENTES(CODICOMPO%, CODICONJU%, NUMSERIE$, NECESIDAD, Optional OK%)
        '
        'SINTESIS
        'RECORRER EL RESERVA SEGUN ORDEN DE FABRICACION
        'OBTENER EL CODIGO , CANTIDAD UNITARIA, Y EL LOTE
        'OBTENER EL STOCK SEGUN EL DEPOSITO DEL ARMADOR, TR, CI%,O.F
        'SI LA CANTIDAD RESERVADA ES <= AL STOCK EN EL DEPOSITO DEL ARMADOR
        'TRANSFERIR DEL ARMADOR AL DEPOSITO DESTINO
        '******************************************
        
        'PARAMETRO PARA PASAR A RUTINA MOVISTOK
        OK% = 1 ' TRANSFERENCIA OK
        '
        FE% = HOY(HOS$)
        CMOV$ = "TR"
        NVM$ = TRIM$(ProNroComprobante("TR"))
        NRO& = XVALO(NVM$)
        DOPRI$ = "TR-" + NVM$
        DEPARMAD% = XVALO(CONTROL("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
        DEPFABRI% = CONTROL("", "DEPODEST")  ' DEPOSITO DE DESTINO
        CONDI$ = "Cod_Inte = " & CODICONJU% & " and NUMEROSERIE = '" & NUMSERIE$ & "'"
        NUMEORF$ = VALOBADA("TANUMSE", "IdSubOr", CONDI$, "NOMESS")
        DOSE$ = TRIM$(NUMEORF$) ' NUMERO DE ORDEN DE FABRICACION
        NUARMA = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IdSubOr = '" & NUMEORF$ & "'", "NOMESS"))
        DARMA$ = ECOARCH$("ARMADORS", MKI$(Int(NUARMA)))
        REFE32$ = "C.F. -Transf.- De Armador A Taller" '+ DARMA$
        TRANFERIDO% = 0
        '
        SQLX$ = "SELECT * FROM RESERVA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMEORF$ & "'"
        SQLX$ = SQLX$ + " AND Cod_Inte = " & CODICOMPO%
        Dim RESTMP As New ADODB.Recordset
        Set RESTMP = READSET(SQLX$)
        i& = 0
        With RESTMP
          Do While Not .EOF
             ' LEE LOS DATOS DEL COMPONENTE EN LA TABLA DE RESERVAS
             CANTIUNIT = XVALO(!USOUNIT)
             LOTEASIG$ = SAFETEXT$(!idenlote)
             CICOMP% = XVALO(!COD_INTE)

             ' OBTENGO LA CANTIDAD ENTRADA Y SALIDA DEL COMPONENTE EN EL DEPOSITO DEL ARMADOR
             Condi1$ = "Cod_Inte = " & CICOMP% & " AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & DOSE$ & "' AND CODIMOVI = 'TR'"
             If TRIM$(LOTEASIG$) <> "" Then Condi1$ = Condi1$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
             STOCK32_ENTRADA = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", Condi1$, "NOMESS"))
             STOCK32_SALIDA = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", Condi1$, "NOMESS"))
             '
'             ' SI EL STOCK32_ENTRADA ES MENOR A CERO, EL COMPONENTE NO FUE TRANSFERIDO
'             If STOCK32_ENTRADA < FRACMIN(CICOMP%) Then ' VERIFICA SI EL MATERIAL FUE TRANSFERIDO
'                OK% = -1
'                Exit Sub
'             End If
             
             ' VERIFICA QUE EL SALDO DEL LOTE EN EL DEPOSITO DEL ARMADOR NO SEA NEGATIVO
             STOCK32 = STOCK32_ENTRADA - STOCK32_SALIDA
             If STOCK32 < FRACMIN(CICOMP%) Then ' STOCK NEGATIVO EN EL DEPOSITO DEL ARMADOR
                OK% = -1
                Exit Sub
             End If
             '
             ' VERIFICA CUANTO PUEDE TRANSFERIR. EL MINIMO ENTRE LA NECESIDAD Y EL STOCK DEL DEPOSITO DEL ARMADOR
             If NECESIDAD <= STOCK32 Then
                CANTIUNIT = NECESIDAD
             Else
                OK% = -1            ' DEVUELVE -1 PORQUE NO ALCANZA PARA CERRAR LA OF
                CANTIUNIT = STOCK32 ' SI NO ALCANZA TRANSFIERE EL MAXIMO DISP. EN EL 32
             End If
             '
             ' VERIFICA SI EXISTE UN TRANSFERENCIA DE SALIDA PARA EL COMPONENTE. EN CASO AFIRMATIVO SOBREESCRIBE LA CANTIDAD AGREGANDO
             ' EL CONSUMO UNITARIO NECESARIO PARA CERRA LA OF.
             ' SI NO HAY UNA TRANSFERENCIA DE SALIDA PREVIA, AGREGA UN NUEVO REGISTRO CORRESPONDIENTE A DICHO MOVIMIENTO POR LA
             ' CANTIDAD SEGUN EL PARAMETRO NECESIDAD.
             If STOCK32_SALIDA > FRACMIN(CICOMP%) Then
                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CICOMP%
                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPARMAD%
                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
                CONDI5$ = CONDI5$ & " AND CANTSALID > 0"
                CANYACO = CANTIUNIT + STOCK32_SALIDA
                Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI5$, CANYACO, RAFE&, "NOMESS")
             Else
                Call MOVISTOK(FE%, CICOMP%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPARMAD%, CANTIUNIT * (-1))
             End If
             '
             NUORIT% = NUORIT% + 1
             '
15           ' VERIFICA SI EXISTE UN TRANSFERENCIA DE ENTRADA PARA EL COMPONENTE. EN CASO AFIRMATIVO SOBREESCRIBE LA CANTIDAD AGREGANDO
             ' EL CONSUMO UNITARIO NECESARIO PARA CERRA LA OF. O SEA, EL PARAMETRO NECESIDAD.
             ' SI NO HAY UNA TRANSFERENCIA DE ENTRADA PREVIA, AGREGA UN NUEVO REGISTRO CORRESPONDIENTE A DICHO MOVIMIENTO POR LA
             ' CANTIDAD SEGUN EL PARAMETRO NECESIDAD.
             CONDI2$ = "Cod_Inte = " & CICOMP% & " AND DepoMovi = " & DEPFABRI% & " And DoSecCor = '" & DOSE$ & "' AND CODIMOVI = 'TR'"
             If TRIM$(LOTEASIG$) <> "" Then CONDI2$ = CONDI2$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
             STOCK32_ENTRADA = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI2$, "NOMESS"))
             
             If STOCK32_ENTRADA > FRACMIN(CICOMP%) Then
                CANYACO = STOCK32_ENTRADA + CANTIUNIT
                Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI2$, CANYACO, RAFE&, "NOMESS")
             Else
                Call MOVISTOK(FE%, CICOMP%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPFABRI%, CANTIUNIT)
             End If
             '
25           .MoveNext
          Loop
        End With
        RESTMP.Close
        Set RESTMP = Nothing
        
        
'        With RESTMP
'          Do While Not .EOF
'             If TRANFERIDO% > 0 Then Exit Do ' SI YA TRANFIRIO UNA VEZ SE VA
'             I& = I& + 1
'             'Call TRACE(24, I&, FIN&)
'             'RECORRO EL RESERVA
'             CANTIUNIT = XVALO(!USOUNIT)
'             LOTEASIG$ = SAFETEXT$(!IDENLOTE)
'             CICOMP% = XVALO(!COD_INTE)
'
'             'OBTENGO EL STOCKS SEGUN O.F. , CODIGO, DEPOSITO Y CODIMOVI = TR
'             CONDI1$ = "Cod_Inte = " & CICOMP% & " AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & DOSE$ & "' AND CODIMOVI = 'TR'"
'             STOCK32 = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE - CANTSALID)", CONDI1$, "NOMESS"))
'             '
'             If CANTIUNIT <= STOCK32 Then
'                'SE PUEDE DAR QUE LA CANTIDAD NECESARIA SEA MAYOR AL USO CUANDO EL SALDO
'                'DEL DEPOSITO 54 ES NEGATIVO, EN ESTE CASO VA A TRANSFERIR PARA CUBRIR
'                If NECESIDAD <= STOCK32 Then
'                   CANTIUNIT = NECESIDAD
'                Else
'                   CANTIUNIT = STOCK32 ' SI NO ALCANZA TRANSFIERE EL MAXIMO DISP. EN EL 32
'                End If
'                'TRANSFIERE DEL 32 AL 54
'                '
'                TRANFERIDO% = 1
'                '
'                ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
'                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
'                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CICOMP%
'                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
'                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPARMAD%
'                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
'                CONDI5$ = CONDI5$ & " AND CANTSALID > 0"
'                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
'                If QTRX& > 0 Then
'                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI5$, "NOMESS"))
'                   Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI5$, CANYACO, RAFE&, "NOMESS")
'                   GoTo 15      ' LO AGREGO AL MOVIMIENTO PREVIOS
'                End If
'                '
'                NUORIT% = NUORIT% + 1
'                Call MOVISTOK(FE%, CICOMP%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPARMAD%, CANTIUNIT * (-1))
'                '
'15              ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
'                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
'                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CICOMP%
'                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
'                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPFABRI%
'                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
'                CONDI5$ = CONDI5$ & " AND CANTINGRE > 0"
'                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
'                If QTRX& > 0 Then
'                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI5$, "NOMESS"))
'                   REFEXX$ = SAFETEXT$(VALOBADA("MOVISTO", "ReferCor", CONDI5$, "NOMESS")) '+ "," + NUMEBORE$
'                   Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI5$, CANYACO, RAFE&, "NOMESS")
'                   Call ACTUREGISTRO("MOVISTO", "ReferCor", CONDI5$, REFEXX$, RAFE&, "NOMESS")
'                   GoTo 25      ' LO AGREGO AL MOVIMIENTO PREVIOS
'                End If
'                '
'                Call MOVISTOK(FE%, CICOMP%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPFABRI%, CANTIUNIT)
'             End If
'             '
'25           .MoveNext
'          Loop
'        End With
'        RESTMP.Close
'        Set RESTMP = Nothing
        
End Sub
                  


Private Sub Command23_Click()
    NUSERIE$ = TRIM$(Text3)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       Exit Sub
    End If
    Call PRINTMEDI(0) '  0 IMPRESION SIMPLE --- UNO GRABACION
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPRCF.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command4_Click()
   
   If XVALO(VALOBADA("TANUMSE", "NumeroSerie", "NumeroSerie ='" & TRIM$(Text3) & "'")) > 0 Then
     Call VERIMPENSAYOS ' SI EL NUMERO DE SERIE ES VALIDO MUESTRA LA LISTA DE SELECCION
   End If
End Sub

Private Sub Command5_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Label101 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Command6_Click()
   '\\\OT-08-0929-OD-06/10/08///
    If Image3.Visible = False Then
      Call CommSel("SELFECHA", Command6, Text2)
    Else
      Command7.SetFocus
    End If
    '\\\OT-08-0929-OD-FIN///
End Sub

Private Sub Command7_Click()
'
Call Text3_DblClick
'
End Sub

Private Sub Command8_Click()
   If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
     Call SELECHO("TADEPOSI")
     If VALACT1$("TADEPOSI") <> "" Then
       Label103 = TRIM$(VALACT1$("TADEPOSI"))
     End If
   End If
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    Call COPYSTRU("INDIASIE", "INCONSQL")
    Set BaDaConsul = OpenDatabase(App.Path + "\SQLCTRL.MDB")
    '
    FORNAME$ = TRIM$(UCase$(Name))
    
    strSQL = "SELECT COD_QUERY FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    MsgBox strSQL
    MsgBox App.Path + "\SQLCTRL.MDB"
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para este Formulario.")
       GoTo 99
    End If
    Do While Not MIRS.EOF
       COCU$ = MIRS!COD_QUERY
       'DECU$ = ECOARCH$("INCONSQL", COCU$)
       CONDI$ = "COD_QUERY = '" & COCU$ & "'"
       DECU$ = VALOTACO("SQL_QUERIES", "DES_QUERY", CONDI$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       JJ& = JJ& + 1
       Call GRAREG("!INDIASIE", TTXX$, JJ&)
    MIRS.MoveNext
    Loop
    MIRS.Close
    Call SETULTREG("!INDIASIE", JJ&)
    Call FRESHECHO("!INDIASIE")
    '
10  RSEL$ = REGSEL$("!INDIASIE/Consultas y Reportes Preconfigurados")
    If Len(RSEL$) > 4 Then
       MODUCONE$ = "?" + TRIM$(Left$(RSEL$, 16))
       DECONEC$ = TRIM$(Mid$(RSEL$, 17, 48))
       Call CONECRUN(MODUCONE$, DECONEC$)
       GoTo 10
    End If
    '
99  Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    HOII% = HOY(HOS$)
    '
    'CREA LOS CAMPOS DE VALORES ADMISIBLES PARA LAS MEDICIONES
    Call CREARCAMPOS_CF_MEDICION
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       Openforms = DoEvents
    End If
    Call BLANFORMU
    '
    FF% = HOY(HO$)
    Text2.TEXT = HO$
    Label83 = USERNAME$
    Text6.TEXT = USERNAME$          ' asigna nombre de usuario alcasillero correspondietne al tecnico
    '
    ' depositos de entrada y consumo
    Label101 = CONTROL("", "DEPOCIE")
    Label103 = CONTROL("", "DEPODEST")
    '
End Sub

Private Sub Image1_DblClick()
   If DERACCE("MODENS70", "Modificacion de Informe de Ensayo") > 1 Then
        Image3.Visible = True
        Image1.Visible = False
        Command2.Enabled = False
        Command23.Enabled = True '\\\OT-08-1064-OD-21/11/08-MEDITEA- HABILITA LA IMPRESORA
        FEC = VALOBADA("TANUMSE", "FECH_ENSA", "NumeroSerie ='" & Text3 & "'")
        FE% = CVINT(FEC) ' SI VUELVE A ESTAR EL  CANDADO VISIBLE PONGO
        Text2 = FECHATEX$(FE%) ' LA FECHA QUE TIENE EL ENSAYO GRABADA
   End If
End Sub

Private Sub Image3_DblClick()

   If DERACCE("MODCONFI", "Modificacion  de Registro de Control Final") > 1 Then
        CONDI$ = "NUMEROSERIE = '" & TRIM$(Text3) & "' AND (NUMREMITO <> '' OR NUMFACTURA <> '') "
        CANTREG& = CantRegistros&("TANUMSE", CONDI$)
        If CANTREG& > 0 Then
           Call COMUNI("Imposible Modificar Informe de Control Final\Equipo Facturado y/o Remitido")
           Exit Sub
        End If
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False
        HOII% = HOY(HOS$)
        Text2.TEXT = HOS$ ' PONGO LA FECHA ACTUAL
   End If
End Sub

Private Sub Label101_Change()
   Label108 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label101)))
End Sub

Private Sub Label101_DblClick()
   Call Command5_Click
End Sub

Private Sub Label103_Change()
   Label102 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label103)))
End Sub

Private Sub Label103_DblClick()
   Call Command8_Click
End Sub

Private Sub Label109_Click(Index As Integer)
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label109(Index) = ECOARCH$("INSTRUME", Chr$(CODINSTRU%))

End Sub

Private Sub Label15_Change()
    Label101 = ""
    Label103 = ""
    NORFAX$ = Label15.Caption
    If NORFAX$ <> "" Then
       'CARGA DEPOSITOS SEGUN CONFIGURACION O SI TIENE ASIGNADO ARMADOR EXTERNO
       Label101 = CONTROL("", "DEPOCIE")
       Label103 = CONTROL("", "DEPODEST")
       'LO SIGUIENTE SE COMENTA PORQUE LOS COMPONENTES TERMINAN QUEDANDO TODOS EN
       'EL DEPOSITO TALLER
       'If ValidaArmador%(NORFAX$) > 0 Then ' CASO ARMADOR EXTERNO
       '   DEPARMA% = XVALO(CONTROL("", "DEPOARMA"))
       '   Label103 = DEPARMA%
       'End If
    End If
End Sub

Private Sub Label15_DblClick()
    If Label15 <> "" Then
     OPPX% = COMALTER%("Elija Opción de Informe:\\Solo Faltantes\Reservas Pendientes Transferidas ")
     If OPPX% < 1 Or OPPX% > 2 Then Exit Sub
     SOLOFALTANTE% = 0 'Reservas Pendientes Transferidas
     If OPPX% = 1 Then SOLOFALTANTE% = 1
      Call INFOCONSU(Label15, SOLOFALTANTE%)
    End If
End Sub

Private Sub Label8_DblClick()
   If Image3.Visible = True Then Exit Sub '\\\OT-08-1074-OD-24/11/08-MEDITEA-si candado visible sale de la rutina
   If DERACCE("CAMBITEN", "Cambio de Tension") > 1 Then
     Text8.Visible = True 'muestra el text de tension y tapa al label
   End If
End Sub

Private Sub Lis1_DblClick(Index As Integer)
  If Image3.Visible = True Then Command7.SetFocus: Exit Sub

   '\\\OT-08-0933-OD-06/10/08///
    If Lis1(Index) = "" Then
       Lis1(Index) = "OK"
       Lis2(Index) = ""
     Else
       Lis1(Index) = ""
    End If
   '\\\OT-08-0933-OD-FIN///
   
   Call VeredGeneral '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN
End Sub

Private Sub Lis2_DblClick(Index As Integer)
   '\\\OT-08-0933-OD-06/10/08///
    If Image3.Visible = True Then Command7.SetFocus: Exit Sub

    If Lis2(Index) = "" Then
       Lis2(Index) = "NC"
       Lis1(Index) = ""
     Else
       Lis2(Index) = ""
    End If
   '\\\OT-08-0933-OD-FIN///
    Call VeredGeneral '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR OK Y NC PARA ANOTAR EL DICTAMEN

End Sub

Private Sub Text1_GotFocus()
  If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub Text2_DblClick()
   Call Command6_Click
End Sub

Private Sub Text2_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub Text3_Change()
   '\\\OT-08-0933-OD-06/10/08///+
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Enabled = False
   Image1.Visible = True
   Command23.Enabled = True
   '
   TXTRAZA_REGICON07$ = ""
   Call LIMPIAINFORMACION
   '
End Sub

Private Sub Text3_DblClick()
   '.
    NUSERIE$ = ListaNumerosSerie
    If NUSERIE$ <> "" Then
      Text3 = NUSERIE$
      Call Text3_KeyPress(13)
      Label93 = TRIM(RESP_ZELE_VECT(4))
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       NUSERIE$ = TRIM$(Text3)
       '
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 1 = " & Now & vbCrLf
       If NUSERIE$ = "" Then Exit Sub
       '
       CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
       If CANTRE& < 1 Then
           Call MENSERR(24, "Número de Serie Incorrecto")
           Exit Sub
       End If
       '
       CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'"
       If CantRegistros&("TANUMSE", CONDI$, "NOMESS") < 1 Then Exit Sub
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 2 = " & Now & vbCrLf
       UBICACI$ = TRIM$(VALOBADA("TANUMSE", "UBICACION", CONDI$, "NOMESS"))
       ORFA$ = TRIM$(VALOBADA("TANUMSE", "IDSUBOR", CONDI$, "NOMESS"))
       ARMA% = XVALO(VALOBADA("ORDEFABR", "Armador", "IDSUBOR='" & ORFA$ & "'", "NOMESS"))
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 3 = " & Now & vbCrLf
       If ARMA% > 0 And UBICACI$ <> "Planta Meditea" Then
           Call COMUNI("Imposible Controlar Este Equipo\No se Encuentra en Planta")
           Text3 = ""
           Exit Sub
       End If
       If InStr(NUSERIE$, "'") < 1 Then
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 4 = " & Now & vbCrLf
           SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
           Dim TNMSETMP As ADODB.Recordset
           'Set TNMSETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
           Set TNMSETMP = READSET(SQLX$)
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 5 = " & Now & vbCrLf
           With TNMSETMP
                If Not .EOF Then
                    '
                    CI% = XVALO(!COD_INTE)
                    Label15 = SAFETEXT$(!IDSUBOR)
                    Label7 = CODEXT$(CI%)
                    Label13 = DESCRIT0$(CI%)
                    Text4 = SAFETEXT$(!Obse_Ensa)
                    Text5 = SAFETEXT$(!Dic_Ensa)
                    Text6 = SAFETEXT$(!Tec_Ensa)
                    If Text6.TEXT = "" Then Text6.TEXT = USERNAME
                    Text7 = SAFETEXT$(!Aext_Ensa)
                    Label8 = SAFETEXT$(!TEN_ENSA)
                    Text8 = XVALO(!TEN_ENSA)
                    Text1 = SAFETEXT$(!AINT_Ensa) '\\\OT-08-1068-OD-21/11/08-MEDITEA-LEVANTA EL ARMADOR INTERNO
                    fecccha$ = SAFETEXT$(!FECH_ENSA) ' si no tiene fecha pone la de hoy
                    If fecccha$ <> "" Then fff% = CVINT(fecccha$): Text2 = FECHATEX(fff%)
                    'si tiene fecha le damos formato "dd/mm/aa"
                    If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Text2.TEXT = HO$
                    Label109(0) = SAFETEXT$(!INSTRUMENTO1)
                    Label109(1) = SAFETEXT$(!INSTRUMENTO2)
                    Label109(2) = SAFETEXT$(!INSTRUMENTO3)
                    Label109(3) = SAFETEXT$(!INSTRUMENTO4)
                    GoSub CargaRestDatos
                Else
                End If
          End With
          'Guarda el cod int, lo compara con el ingresado
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 6 = " & Now & vbCrLf
         
          
CargaRestDatos:
           '
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 7 = " & Now & vbCrLf
           SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM CONTROLFIN WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
           SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
           SQLZ$ = SQLZ$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"
           Dim MAXREVTMP As ADODB.Recordset
           'Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
           Set MAXREVTMP = READSET(SQLZ$)
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 8 = " & Now & vbCrLf
           ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
           MAXREVTMP.Close
           Set MAXREVTMP = Nothing
           Label85 = ULTREVIS%
           '
           SQLX$ = "SELECT NumeroOrden,Resultado FROM CONTROLFIN WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
           SQLX$ = SQLX$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"
           SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
    
           Dim CTRLFINTMP As ADODB.Recordset
           'Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
           Set CTRLFINTMP = READSET(SQLX$)
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 9 = " & Now & vbCrLf
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
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 10 = " & Now & vbCrLf
           J& = 0
           SQLX$ = "SELECT MEDA,MEDB,MEDC FROM MEDIFIN WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
           SQLX$ = SQLX$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"
           SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
           SQLX$ = SQLX$ + " ORDER BY CHN "
    
           Dim MEDIFINTMP As ADODB.Recordset
           'Set MEDIFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
           Set MEDIFINTMP = READSET(SQLX$)
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 11 = " & Now & vbCrLf
           With MEDIFINTMP
            If Not .EOF Then
             Call bloquear
           
             Do While Not .EOF
               If J& > 11 Then Exit Do
               TextA(J&) = FORMATNUM$(XVALO(!MEDA), "F7.2")  '\\\OT--08-1060-20/11/08///MEDITEA MODIFICAR EL FORMATEO A F7.2
               TextB(J&) = FORMATNUM$(XVALO(!MEDB), "F7.2")
               TextC(J&) = FORMATNUM$(XVALO(!MEDC), "F7.2")
               J& = J& + 1
             .MoveNext
             Loop
            End If
           End With
       End If
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 12 = " & Now & vbCrLf
       'Mostrar el valor de CantApro para ese producto, informando la candidad de OK para su aprobacion
       CONDI$ = "Codint =" & CI%
       Label90 = VALOBADA("MASTER", "CantApro", CONDI$, "NOMESS")
       '\\\OT-08-0933-OD-FIN///
       '
       'CARGA LOS VALORES ADMISIBLES PARA LAS MEDICIONES
    '   CONDI$ = "MARBORRA = 0 AND COD_INTE =" & CI%
    '   Label82(0) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MINIMO_1", CONDI$, "NOMESS")), "F7.2"))
    '   Label82(1) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MAXIMO_1", CONDI$, "NOMESS")), "F7.2"))
    '   Label96(0) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MINIMO_2", CONDI$, "NOMESS")), "F7.2"))
    '   Label96(1) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MAXIMO_2", CONDI$, "NOMESS")), "F7.2"))
    '   Label100(0) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MINIMO_3", CONDI$, "NOMESS")), "F7.2"))
    '   Label100(1) = TRIM$(FORMATNUM$(XVALO(VALOBADA("CF_MEDICION", "MAXIMO_3", CONDI$, "NOMESS")), "F7.2"))
       '
       SQLX$ = "SELECT * FROM CF_MEDICION WITH(NOLOCK) WHERE COD_INTE =" & CI% & " AND MARBORRA = 0"
       Set rst = READSET(SQLX$)
    TXTRAZA_REGICON07$ = TXTRAZA_REGICON07$ & "PUNTO 13 = " & Now & vbCrLf
       With rst
            If Not (.EOF And .BOF) Then
                Label82(0) = TRIM$(FORMATNUM$(XVALO(!MINIMO_1), "F7.2"))
                Label82(1) = TRIM$(FORMATNUM$(XVALO(!MAXIMO_1), "F7.2"))
                Label96(0) = TRIM$(FORMATNUM$(XVALO(!MINIMO_2), "F7.2"))
                Label96(1) = TRIM$(FORMATNUM$(XVALO(!MAXIMO_2), "F7.2"))
                Label100(0) = TRIM$(FORMATNUM$(XVALO(!MINIMO_3), "F7.2"))
                Label100(1) = TRIM$(FORMATNUM$(XVALO(!MAXIMO_3), "F7.2"))
            End If
       End With
       On Error Resume Next
       rst.Close
       Set rst = Nothing
       On Error GoTo 0
    End If
End Sub

Private Sub Text4_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub Text5_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub Text6_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus
End Sub


Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub Text8_GotFocus()
   If Image3.Visible = True Then Command7.SetFocus '\\\OT-08-1074-OD-24/11/08-MEDITEA-si candado visible pone el foco en command7
End Sub

Private Sub TextA_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub TextA_LostFocus(Index As Integer)
   '\\\OT-08-0933-OD-06/10/08///
  VALOR1 = ROUND#(XVALO(TextA(Index)), 2)  '\\\OT--08-1060-20/11/08///MEDITEA
'  MINIMOX = XVALO(Label82(0))
'  MAXIMOX = XVALO(Label82(1))
  If XVALO(VALOR1) < 10000 Then 'AMPLIAR EL INGRESO PERMITIDO a 9999.99
      TextA(Index) = FORMATNUM(TextA(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
  End If
   '\\\OT-08-0933-OD-FIN///
End Sub
Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label85)
    If N% > 0 Then
        REVIS% = 0
        NUSERIE$ = TRIM$(Text3)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM CONTROLFIN WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        SQLZ$ = SQLZ$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB')"

        Dim MAXREVTMP As ADODB.Recordset
        'Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        Set MAXREVTMP = READSET(SQLZ$)
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    TIPODOC$ = "Registro de Control Final" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Text3) & "-" & REVIS%
    '
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Text2) 'FECHA
    CODPARAM$(2) = "NRO-AENT": DOCPARAM$(2) = Label15 'o.f
    CODPARAM$(3) = "TALLE": DOCPARAM$(3) = TRIM$(Text6) 'TECNICO
    CODPARAM$(4) = "COLOR": DOCPARAM$(4) = TRIM$(Text7) 'A.EXTERNO
    CODPARAM$(5) = "REF-MAT1":  DOCPARAM$(5) = TRIM$(Text3.TEXT) 'NRO SERIE
    CODPARAM$(6) = "DES-MAT": DOCPARAM$(6) = Label13.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(7) = "REF-MAT2":  DOCPARAM$(7) = Text5.TEXT  'DICTAMEN
    CODPARAM$(8) = "OBSERVA":  DOCPARAM$(8) = Text4.TEXT  'OBSERVACIONES
       NUSER1$ = TRIM$(Text3)
       NORFA$ = TRIM$(Label15)
       '
       CONDI$ = " IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NUSER1$ & "'"
       CIXI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", CONDI$, "NOMESS"))
       TENSION1% = XVALO(VALOBADA("TANUMSE", "Ten_Ensa", CONDI$, "NOMESS"))
       TENSION2$ = "220"
       If TENSION1% < 150 Then
         TENSION2$ = TALLE$(CIXI%)
       Else
         TENSION2$ = COLOR$(CIXI%)
       End If
    CODPARAM$(9) = "N-HPROF":  DOCPARAM$(9) = TENSION2$ 'ANTES TEXT8'-TENSION
    CODPARAM$(10) = "REF-MAT3": DOCPARAM$(10) = Text1.TEXT
    'SI ES POR GRABACION IMPRIME EL NUMERO DE CONTROL
    NUCONTR$ = ""
    If N% > 0 Then
      NUCONTR$ = FORMATNUM$(NUMECONTROL&, "F8.0")
    End If
    CODPARAM$(11) = "N-OCOMPR": DOCPARAM$(11) = NUCONTR$
    '
    'INSTRUMENTOS
    CODPARAM$(12) = "REF-MAT4": DOCPARAM$(12) = TRIM$(Label109(0))
    CODPARAM$(13) = "REF-MAT5": DOCPARAM$(13) = TRIM$(Label109(1))
    CODPARAM$(14) = "REF-MAT6": DOCPARAM$(14) = TRIM$(Label109(2))
    CODPARAM$(15) = "REF-CONJ": DOCPARAM$(15) = TRIM$(Label109(3))
    '
    CAPARDOC% = 15
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
        TETABU1$(2, CAITAB1%) = TRIM$(TextB(U&))     'MEDICIONES B
        TETABU1$(3, CAITAB1%) = TRIM$(TextC(U&))     'MEDICIONES C
    Next U&
    CAITAB1% = 0
    For J& = 0 To 14
        CAITAB1% = CAITAB1% + 1
        TETABU1$(4, CAITAB1%) = TRIM$(Lis1(J&))   ' NUMERO
        TETABU1$(5, CAITAB1%) = TRIM$(Lis2(J&))    'CONTROL
        TETABU1$(6, CAITAB1%) = TRIM$(Lis1(J& + 15)) ' NUMERO
        TETABU1$(7, CAITAB1%) = TRIM$(Lis2(J& + 15))  ' CONTROL
    Next J&
    '
    FORIPRE$ = TRIM$(CONTROL$("", "RECONFI"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    i& = 0
    For i& = 0 To 29
        Lis1(i&) = ""
        Lis2(i&) = ""
    Next i&
    i& = 0
    For i& = 0 To 11
        TextA(i&) = ""
        TextB(i&) = ""
        TextC(i&) = ""
    Next i&
    Label15 = ""
    Text3 = ""
    Label7 = ""
    Label13 = ""
    Text5 = ""
    Label8 = ""
    Text1 = "" '\\\OT-08-1068-OD-21/11/08-MEDITEA-LIMPIA EL ARMADOR INTERNO
    Text4 = ""
    Text6 = ""
    Text7 = ""
    Text2 = FECHATEX$(HOY(HO$))
    Label85 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA- BORRAR CAPTION DEL LABEL
    Text8.Visible = False '\\\OT-08-1074-OD-24/11/08-MEDITEA-oculto el text8
    For i& = 0 To 1
      Label82(i&) = ""
      Label96(i&) = ""
      Label100(i&) = ""
    Next i&
    For i& = 0 To 3
      Label109(i&) = ""
    Next i&

End Sub

Private Sub TextB_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub TextB_LostFocus(Index As Integer)
  VALOR1 = ROUND#(XVALO(TextB(Index)), 2)  '\\\OT--08-1060-20/11/08///MEDITEA

  If XVALO(VALOR1) < 10000 Then  'AMPLIAR EL INGRESO PERMITIDO a 9999.99
    TextB(Index) = FORMATNUM(TextB(Index), "F7.2")
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    TextB(Index).SetFocus
  End If

End Sub

Private Sub TextC_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command7.SetFocus
End Sub

Private Sub TextC_LostFocus(Index As Integer)
  VALOR1 = ROUND#(XVALO(TextC(Index)), 2)  '\\\OT--08-1060-20/11/08///MEDITEA
  '
  If XVALO(VALOR1) < 10000 Then
    TextC(Index) = FORMATNUM(TextC(Index), "F7.2") 'AMPLIAR EL INGRESO PERMITIDO a 9999.99
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 10000.\Debe Corregir Para Poder Continuar.")
    TextC(Index).SetFocus
  End If
End Sub
Sub LIMPIAINFORMACION()
    Label8 = ""
    Label13 = ""
    Label15 = ""
    Label7 = ""
    Text1 = "" '\\\OT-08-1068-OD-21/11/08-MEDITEA-LIMPIA EL ARMADOR INTERNO
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Label90 = ""
    Label93 = ""
    i& = 0
    For i& = 0 To 29
        Lis1(i&) = ""
        Lis2(i&) = ""
    Next i&
    i& = 0
    For i& = 0 To 11
        TextA(i&) = ""
        TextB(i&) = ""
        TextC(i&) = ""
    Next i&
    Label85 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA- BORRAR CAPTION DEL LABEL
    Text8.Visible = False '\\\OT-08-1074-OD-24/11/08-MEDITEA-oculto el text8
    For i& = 0 To 1
      Label82(i&) = ""
      Label96(i&) = ""
      Label100(i&) = ""
    Next i&
    For i& = 0 To 3
      Label109(i&) = ""
    Next i&

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
  'Recorre Lis2 si encuentra un NC muestra RECHA
  For i& = 0 To 29
    If Lis2(i&) = "NC" Then
      Text5 = "RECHA"
      Exit Sub
    End If
  Next i&
  Text5 = ""
  
  'Recorre Lis2 si encuentra un NC, lo acumula
  For i& = 0 To 29
    If Lis1(i&) = "OK" Then
      OK% = OK% + 1
    End If
  Next i&
  
  
  'Guardar valor del Label90, y si es igual al valor de OK%, muestra APRO sino RECHA
  CANPRO% = XVALO(Label90)
  If CANPRO% > 0 Then
    If OK% >= CANPRO% Then
      Text5 = "APRO"
    Else
      Text5 = "RECHA"
    End If
  End If
End Sub
Sub VERIMPENSAYOS()
   Command4.Enabled = False
   '
   Call SETULTREG("!INDIENSA", 0)
   '
   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM CONTROLFIN WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Text3 & "'"
   SQLX$ = SQLX$ + " AND (ORIGEN IS NULL OR ORIGEN = 'FAB' OR ORIGEN='')"

   Dim CTRLFINTMP As ADODB.Recordset
   'Set CTRLFINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set CTRLFINTMP = READSET(SQLX$)
   '
   With CTRLFINTMP 'COMPLETO CON LOS DATOS DE CONTROL FIN
     U& = 0
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!FECH_ENSA)
       NUREV$ = "Registro de Control Final " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENSA", TX$, U&)
       .MoveNext
     Loop
   End With
   CTRLFINTMP.Close
   Set CTRLFINTMP = Nothing
   '
   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM ENSAR070 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Text3 & "'"
   Dim ENSR070TMP As ADODB.Recordset
   'Set ENSR070TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set ENSR070TMP = READSET(SQLX$)
   '
   With ENSR070TMP 'COMPLETO CON LOS DATOS DE ENSAR070
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!FECH_ENSA)
       NUREV$ = "Puesta a Tierra " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENSA", TX$, U&)
       .MoveNext
     Loop
   End With
   ENSR070TMP.Close
   Set ENSR070TMP = Nothing
   '
   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM ENSAR072 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Text3 & "'"
   Dim ENSR072TMP As ADODB.Recordset
   'Set ENSR072TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set ENSR072TMP = READSET(SQLX$)
   '
   With ENSR072TMP 'COMPLETO CON LOS DATOS DE ENSAR072
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!FECH_ENSA)
       NUREV$ = "Fuga a Tierra " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENSA", TX$, U&)
       .MoveNext
     Loop
   End With
   ENSR072TMP.Close
   Set ENSR072TMP = Nothing
   '
   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM ENSAR074 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Text3 & "'"
   Dim ENSR074TMP As ADODB.Recordset
   'Set ENSR074TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set ENSR074TMP = READSET(SQLX$)
   '
   With ENSR074TMP   'COMPLETO CON LOS DATOS DE ENSAR074
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!FECH_ENSA)
       NUREV$ = "Registro de Medicion " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENSA", TX$, U&)
       .MoveNext
     Loop
   End With
   ENSR074TMP.Close
   Set ENSR074TMP = Nothing
   '
   SQLX$ = "SELECT DISTINCT NumeroSerie, REVISION, FECH_ENSA FROM SEGELECT WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & Text3 & "'"
   Dim SEGELECTTMP As ADODB.Recordset
   'Set SEGELECTTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   Set SEGELECTTMP = READSET(SQLX$)
   '
   With SEGELECTTMP 'COMPLETO CON LOS DATOS DE SEGELECT
     Do While Not .EOF
       U& = U& + 1
       NUSRIE$ = SAFETEXT$(!NUMEROSERIE)
       RVSON% = XVALO(!Revision)
       FENSA% = CVINT(!FECH_ENSA)
       NUREV$ = "Fuga a Tierra " & NUSRIE$ & "-" & RVSON%
       Call REPLA(TX$, MKI$(FENSA%), 1, 2)
       Call REPLA(TX$, NUREV$, 3, 62)
       Call GRAREG("!INDIENSA", TX$, U&)
       .MoveNext
     Loop
   End With
   SEGELECTTMP.Close
   Set SEGELECTTMP = Nothing

   Call SETULTREG("!INDIENSA", U&)
   Call FRESHECHO("!INDIENSA")
   '
10 If ULTREG&("!INDIENSA") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
     Call SELECHO("!INDIENSA")
   Else
     GoTo 99
   End If
   '
   NUDOCU$ = VALACT2$("!INDIENSA") ' MUESTRO EL FORMULARIO IMPRESO  SELECCIONAD0
   If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
    POS% = InStr(NUDOCU$, TRIM$(Text3)) ' ASIGNO A POS% LA POSICION DESDE DONDE EMPIEZA EL NUMERO DE SEREIE
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
 
Sub ENTRAEQUIPO()
    '
    CAN = 1: NUORIT% = 0: NC% = 0
    DEPX% = XVALO(Label101)   ' DEPOSITO DE DESTINO
    FECHA% = HOY(HOS$)
    CIXI% = CODINT%(Label7)
    LOTE$ = Text3.TEXT        ' NUMERO DE SERIE
    CMOV$ = "OF"
    DOPRI$ = Label15.Caption  ' NUMERO DE ORDEN DE FABRICACION
    DOSEC$ = "IF-" + TRIM$(Mid$(DOPRI$, 4))
    REFE$ = "Fabricación Propia"
    VUNI# = 0
    MONE$ = "$"
    
    Dim obj As New Art_Importados
    CODI$ = Label7.Caption
    Call obj.VeriImportado(CODI$) 'VERIFICA QUE NO SEA IMPORTADO
    
    If XVALO(Label93) > 0 And obj.EsImportado > 0 Then 'SI ES IMPORTADO INGRESA LA MERCADERIA AL DEPOSITO CONFIGURADO
    'TRATAMIENTOS DE EQUIPOS IMPORTADOS
       CAN = CAN * -1 'CAMBIA A NEGATIVO LA CANTIDAD PARA QUE PROVOQUE UNA SALIDA DEL DEPOSITO
       DEPO_RECEP_EQUIPOS_IMPORTADOS% = CONTROL("BOLRECEP", "DEPEQIMP")
       Call MOVISTOK(FECHA%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO_RECEP_EQUIPOS_IMPORTADOS%, CAN)
       CAN = CAN * -1 'LO PASO A POSITIVO
       DEPOIMPORTADO% = XVALO(CONTROL("", "DEPOIMPO"))
       Call MOVISTOK(FECHA%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOIMPORTADO%, CAN)
    'FIN TRATAMIENTO DE EQUIPOS IMPORTADOS
    Else
       'MODO NORMAL
       Call MOVISTOK(FECHA%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
    End If
    '
End Sub

Sub ANOTAENORFA()
    '
    NORFAX$ = TRIM$(Label15.Caption)
    CANTI = "1": NUSERIE$ = TRIM$(Text3)
    TTXYZ$ = Space$(76)
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, "Cierre de O/Fabricación", 11, 56)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call AGREGA_ANOTAORFAB(NORFAX$, TTXYZ$)
    '
    TTXYZ$ = Space$(76)
    Call REPLA(TTXYZ$, " - Cantidad Aprob.= " & CANTI & " - S.Nro: " & NUSERIE$, 11, 56)
    Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
    Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
    '
End Sub

Sub CONSUCOMPO()
    '
    CMOV$ = "CF"
    NORFAX$ = TRIM$(Label15)
    DOPRI$ = NORFAX$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
    FIMATER& = ULTREG&("!MATEROF")
    For U& = 1 To FIMATER&
      Call TRACE(20, U&, FIMATER&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CICOMPO% = CVI(Left$(ZZ$, 2))
      CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
      If TRIM$(UCase$(Unimed$(CICOMPO%))) = "GK" Then CAN = CAN / 1000
         CONDI$ = "CODIMOVI = 'CF' "
         CONDI$ = CONDI$ & " AND COD_INTE = " & CICOMPO%
         CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
         CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
      CANYACO = XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI$, "NOMESS"))
      CANYACO = CANYACO - XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI$, "NOMESS"))
      CAN = CVS(Mid$(ZZ$, 5, 4)) + CANYACO
      Call REPLA(ZZ$, MKS$(CAN), 5, 4)
      Call GRAREG("!MATEROF", ZZ$, U&)
    Next U&
    '
    ' MARCA LOS REGISTROS A BORRAR
    Call CREACAMPO("STOCKS", "MOVISTO", "MarBorra", 3, 2, 1)
    CONDI$ = "CODIMOVI = 'CF' "
    CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
    Call ACTUREGISTRO("MOVISTO", "MarBorra", CONDI$, 1, RAFE&, "NOMESS")
    '
    ' AGREGA EL CONSUMO
    NMS% = NUMAS%
    DP2% = XVALO(Label103)
    NUORIT% = 0
    For U& = 1 To FIMATER&
      Call TRACE(20, U&, FIMATER&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CICOMPO% = CVI(Left$(ZZ$, 2))
      CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
      USOUNI# = CVD(Mid$(ZZ$, 9, 8))
      CODICONJ$ = Label7
      If CICOMPO% > 0 Then
        If CICOMPO% <= NMS% Then
          If Abs(CAN) >= 0.05 Then
            REFE$ = "CONSUMO " + CODICONJ$
            If CAN > 0 Then REFE$ = "RECUPERO " + CODICONJ$
            '
            DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CICOMPO%)
144         NUORIT% = NUORIT% + 1
            '
            ' determina el lote reservado
            LOTE$ = ""
            CONDIX$ = " IdSubOr = '" & NORFAX$ & "' AND Cod_Inte = " & CICOMPO% & ""
            LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
            '
            FECHA% = HOY(HOS$)
            Call MOVISTOK(FECHA%, CICOMPO%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
            CIMOVI%(CICOMPO%) = 1
          End If
        End If
      End If
    Next U&
    Call CIERRARCH("!MATEROF")
    '
    ' BORRA LOS CONSUMOS ANTERIORES
    CONDI$ = "CODIMOVI = 'CF' "
    CONDI$ = CONDI$ + " AND MARBORRA = 1 "
    CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
    Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&, "NOMESS")
    '
    ' cancela las reservas de los componente
    SQLX$ = "SELECT COD_INTE,CANTRES,CANTCONS,SALDO_ENTRE FROM Reserva "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
    Dim RESETEMP As ADODB.Recordset
    Set RESETEMP = New ADODB.Recordset
    RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RESETEMP
       Do While Not .EOF
         CICOMPO% = !COD_INTE
         CARESE# = !CANTRES
           CONDI$ = "CODIMOVI = 'CF' "
           CONDI$ = CONDI$ & " AND COD_INTE = " & CICOMPO%
           CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
           CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
         CANYACO = XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI$, "NOMESS"))
         CANYACO = CANYACO - XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI$, "NOMESS"))
         !CANTCONS = CANYACO
         SALDO = CARESE# - CANYACO
         If SALDO < 0 Then SALDO = 0
         !Saldo_Entre = SALDO
         .Update
       .MoveNext
       Loop
    End With
    RESETEMP.Close
    Set ReserTemp = Nothing
    '
End Sub

Sub CONSUCOMPO_1()
    '
    NORFAX$ = TRIM$(Label15)
    DP2% = XVALO(Label103)
    DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CICOMPO%)
    '
    ARCHIVOX1$ = "MATEROF"
    'EPAC$ = TRIM$(UCase$(EMPREAC$))
    'If InStr(EPAC$, "MEDITEA") > 0 Then ARCHIVOX1$ = "MATEROF9"

    FIMATER& = ULTREG&("!" & ARCHIVOX1$)
    '
    For U& = 1 To FIMATER&
       Call TRACE(20, U&, FIMATER&)
       ZZ$ = REGLEIDO$("!" & ARCHIVOX1$, U&)
       CICOMPO% = CVI(Left$(ZZ$, 2))
       CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
       If TRIM$(UCase$(Unimed$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
       '
       FECHA% = HOY(HOS$)
       Call CONSUMECOMPO(CICOMPO%, NORFAX$, CANCONSUM, FECHA%, DEPX%)
    Next U&
    '
End Sub
Sub CREARCAMPOS_CF_MEDICION()

   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "Cod_inte", 3, 0, 2)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "CODEXTE", 10, 24)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MINIMO_1", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MAXIMO_1", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MINIMO_2", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MAXIMO_2", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MINIMO_3", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MAXIMO_3", 7, 8)
   Call CREACAMPO("EQUIPOS", "CF_MEDICION", "MARBORRA", 3, 0, 1)

   Call CreaIndiceCompuesto("CF_MEDICION", "MARBORRA+CODEXTE")
End Sub



Sub INFOCONSU(NORFAX$, SOLOFALTANTE%)
   'SOLOFALTANTE% = 1 MUESTRA SOLO LOS FALTANTES PARA ESTA ORDEN
   'SOLOFALTANTE% = 0 MUESTRA TODOS LOS COMPONENTES RESERVADOS PENDIENTES TRANSFERIDOS A DEP. 32 Y 54
'   NORFAX$ = "OF-057786-1"
   'VALIDO QUE HAYA INFORMACIION  EN LA RESERVA
   CONDI$ = " IdSubOr = '" & NORFAX$ & "' AND (CANTRES-CANTCONS > 0.05)"
   If CantRegistros&("RESERVA", CONDI$, "NOMESS") < 1 Then
     Call COMUNI("NO SE HAN ENCONTRADO RESERVAS PENDIENTES PARA ESTA ODRDEN DE FABRICACION")
     Exit Sub
   End If
   '------------------------------------------------------------------------------------------
   'TALLERES DE CONSUMO Y ARMADOR
   DEPOTALLER% = XVALO(CONTROL("", "DEPODEST"))
   DEPARMAD% = XVALO(CONTROL("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
   '------------------------------------------------------------------------------------------
   '
   'CARGO LOS ENCABEZADOS DE LA GRILLA Y ADAPTO EL FORMULARIO
   INFO_GRILLA07.Caption = "Situación de Reservas Transferidas ( A Dep.: " & DEPOTALLER% & " o " & DEPARMAD% & " ) Para Orden de Fabricación: " & NORFAX$
   Call CARGA_ENCABEZADO_GRILLA
   '------------------------------------------------------------------------------------------
   '
   'RECORRO LA TABLA RESERVA SEGUN LA ORDEN DE FABRICACION (QUE TRAE LO PENDIENTE)
   REG& = 0
   SQLX$ = "SELECT  COD_INTE,COD_EXTE, IDENLOTE  FROM RESERVA WITH(NOLOCK) WHERE " & CONDI$
   SQLX$ = SQLX$ + " GROUP BY COD_INTE, COD_EXTE, IDENLOTE"
   SQLX$ = SQLX$ + " ORDER BY COD_EXTE, IDENLOTE"
'
   Set RESTMP = READSET(SQLX$)
   With RESTMP
      Do While Not .EOF
        CI% = XVALO(!COD_INTE)
        CODIEX$ = CODEXT$(CI%)
        DESCRI$ = DESCRIT0$(CI%)
        CODIEQUIPO% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IdSubOr = '" & NORFAX$ & "'"))
        CANTIUNIT = USOUNITARIO#(CODIEQUIPO%, CI%)
        '
        TRAZA$ = "NO"
        ESTRA% = ESTRAZABLE%(CI%)
        If ESTRA% > 0 Then TRAZA$ = "SI"
        '
        ENTRO% = 0
        DISPONIBLE_ACUMULDO = 0
        LOTE_ACUMULADO$ = ""
        '
        'BUSCO LO TRANSFERIDO PARA EL CODIGO QUE ESTA LEYENDO DEL RESERVA
        SQLX$ = "SELECT SUM(CANTINGRE) AS TRANSFERIDO, IDENLOTE"
        SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI% & " AND DOSECCOR = '" & TRIM$(NORFAX$) & "'"
        SQLX$ = SQLX$ + " AND (DepoMovi = " & DEPOTALLER% & " OR DepoMovi = " & DEPARMAD% & ") "
        SQLX$ = SQLX$ + " GROUP BY  IDENLOTE "
        '
        'RECORRE EL MOVISTO ACUMULANDO QUE STOCKS HAY DISPONIBLE DEL ARTICULO
        'CONSIDERANDO LO TRANSFERIDO MENOS LO CONSUMIDO , TAMBIEN SE CONSIDERA EL DEPOSITO DE
        'CONSUMO Y EL DE ARMADORES (POR QUE AL CERRAR SI HAY EN EL DE ARMADOR  SE PASA AL DE CONSUMO)
        Set MOVTMP = READSET(SQLX$)
        With MOVTMP
10         If Not .EOF Then
             STOCK_TRANSFERIDO = XVALO(!TRANSFERIDO)
             LOTE$ = TRIM$(SAFETEXT$(!idenlote))
               CONDI$ = "CodiMovi = 'CF' AND Cod_Inte = " & CI% & " AND DOPRICOR = '" & TRIM$(NORFAX$) & "' "
             STOCK_CONSUMIDO = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
             DISPONIBLE = STOCK_TRANSFERIDO - STOCK_CONSUMIDO: If DISPONIBLE < 0.005 Then DISPONIBLE = 0
             '
             DISPONIBLE_ACUMULDO = DISPONIBLE_ACUMULDO + DISPONIBLE
             If LOTE$ <> "" Then
                If LOTE_ACUMULADO$ <> "" Then
                   LOTE_ACUMULADO$ = LOTE_ACUMULADO$ & ";" & LOTE$
                Else
                   LOTE_ACUMULADO$ = LOTE$
                End If
             End If
             
             .MoveNext
             GoTo 10
           'CUANDO ACUMULO EL DISPONIBLE ENTRA ACA Y REVISA SI ES SUFICIENTE SI NO LO PRESENTA
           Else
             If DISPONIBLE_ACUMULDO < CANTIUNIT - FRACMIN(CI%) Or SOLOFALTANTE% < 1 Then
               REG& = REG& + 1
               INFO_GRILLA07.MSF.Rows = REG& + 1
               FALTANTE = CANTIUNIT - DISPONIBLE_ACUMULDO
               If FALTANTE < FRACMIN(CI%) Then FALTANTE = 0
               '
               INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = CODIEX$
               INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = DESCRI$
               INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = FORMATNUM$(CANTIUNIT, "F12." & CNTDC$(CI%))
               INFO_GRILLA07.MSF.TextMatrix(REG&, 4) = TRAZA$
               INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = LOTE_ACUMULADO$
               INFO_GRILLA07.MSF.TextMatrix(REG&, 6) = FORMATNUM$(DISPONIBLE_ACUMULDO, "F12." & CNTDC$(CI%))
               INFO_GRILLA07.MSF.TextMatrix(REG&, 7) = FORMATNUM$(FALTANTE, "F12." & CNTDC$(CI%))
               '
             End If '
           End If

        End With
        MOVTMP.Close
        Set MOVTMP = Nothing
        '
        .MoveNext
      Loop
   End With
   RESTMP.Close
   Set RESTMP = Nothing
   
   If REG& > 0 Then
     INFO_GRILLA07.Show 1
   Else
     Call COMUNI("No se Registran Faltantes Para Este Equipo")
   End If
   
End Sub
Sub CARGA_ENCABEZADO_GRILLA()
   With INFO_GRILLA07
     .MSF.Clear
     .MSF.Cols = 8
     NUCOLU% = 1: .TEPRUEBA = "123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
     NUCOLU% = 2: .TEPRUEBA = "12345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Código"
     NUCOLU% = 3: .TEPRUEBA = "123456789012345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Descripción"
     NUCOLU% = 4: .TEPRUEBA = "1234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Necesidad"
     NUCOLU% = 5: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Trazable"
     NUCOLU% = 6: .TEPRUEBA = "123456789012312345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Lote"
     NUCOLU% = 7: .TEPRUEBA = "1234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Disponible"
     NUCOLU% = 8: .TEPRUEBA = "1234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Faltante"
   End With
   '
   'TAMAÑO DE FRM DE GRILLA
   INFO_GRILLA07.Width = 12800
   INFO_GRILLA07.MSF.Width = 12500
   INFO_GRILLA07.Command8.Left = 10300
   '
End Sub


